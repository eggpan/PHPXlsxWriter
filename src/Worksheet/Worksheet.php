<?php

namespace Eggpan\PHPXlsxWriter\Worksheet;

use Eggpan\PHPXlsxWriter\Cell\Coordinate;
use Eggpan\PHPXlsxWriter\Spreadsheet;
use Eggpan\PHPXlsxWriter\Writer\Xml;
use Eggpan\PHPXlsxWriter\Exceptions\UnknownErrorException;

class Worksheet
{
    protected $spreadsheet;
    protected $name;
    protected $currentColor;
    protected $data = [];
    protected $dimension;
    protected $lastBorder;
    protected $lastBorderPosition;
    protected $lastStyle;
    protected $lastFont;
    protected $lastColor;
    protected $lastFill;
    protected $lastFillEndColor;
    protected $lastFillStartColor;
    protected $mergeCells = [];

    public function __construct(Spreadsheet $spreadsheet, string $name)
    {
        $this->spreadsheet = $spreadsheet;
        $this->name = $name;
    }

    /**
     * 枠線の設定準備をする
     *
     * @return Worksheet
     */
    public function getBorders(): Worksheet
    {
        $this->lastBorder = $this->lastStyle;
        return $this;
    }

    /**
     * 下枠線の設定準備をする
     *
     * @return Worksheet
     */
    public function getBottom(): Worksheet
    {
        $this->lastBorderPosition = 'bottom';
        return $this;
    }

    /**
     * 文字色の設定準備をする
     *
     * @return Worksheet
     */
    public function getColor(): Worksheet
    {
        $this->lastColor = $this->lastFont;
        $this->currentColor = 'color';
        return $this;
    }

    /**
     * 塗りつぶしパターン色の設定準備をする
     *
     * @return Worksheet
     */
    public function getEndColor(): Worksheet
    {
        $this->lastFillEndColor = $this->lastStyle;
        $this->currentColor = 'fillEnd';
        return $this;
    }

    /**
     * 左枠線の設定準備をする
     *
     * @return Worksheet
     */
    public function getLeft(): Worksheet
    {
        $this->lastBorderPosition = 'left';
        return $this;
    }

    /**
     * スタイル取得したセルに対して塗りつぶしの準備をする
     *
     * @return Worksheet
     */
    public function getFill(): Worksheet
    {
        $this->lastFill = $this->lastStyle;
        return $this;
    }

    /**
     * 文字の設定準備をする
     *
     * @return Worksheet
     */
    public function getFont(): Worksheet
    {
        $this->lastFont = $this->lastStyle;
        return $this;
    }

    /**
     * シート名を取得する
     *
     * @return void
     */
    public function getName()
    {
        return $this->name;
    }

    /**
     * 右枠線の設定準備をする
     *
     * @return Worksheet
     */
    public function getRight(): Worksheet
    {
        $this->lastBorderPosition = 'right';
        return $this;
    }

    /**
     * 塗りつぶし色の設定準備をする
     *
     * @return Worksheet
     */
    public function getStartColor(): Worksheet
    {
        $this->lastFillStartColor = $this->lastStyle;
        $this->currentColor = 'fillStart';
        return $this;
    }

    /**
     * 指定したセルの設定準備をする
     *
     * @param string  $cell
     * @return Worksheet
     */
    public function getStyle(string $cell): Worksheet
    {
        list($column, $row) = Coordinate::columnRowIndexFromString($cell);
        return $this->getStyleByColumnAndRow($column, $row);
    }

    /**
     * 指定したセルの設定準備をする
     *
     * @param integer $column
     * @param integer $row
     * @return Worksheet
     */
    public function getStyleByColumnAndRow(int $column, int $row): Worksheet
    {
        $this->lastStyle = [
            'column' => $column,
            'row' => $row,
        ];
        return $this;
    }

    /**
     * 上枠線の設定準備をする
     *
     * @return Worksheet
     */
    public function getTop(): Worksheet
    {
        $this->lastBorderPosition = 'top';
        return $this;
    }

    /**
     * 指定した範囲のセルをマージする
     *
     * @param string $cells
     * @return Worksheet
     */
    public function mergeCells(string $cells): Worksheet
    {
        $this->mergeCells[] = $cells;
        return $this;
    }

    /**
     * 指定した範囲のセルをマージする
     *
     * @param integer $col1
     * @param integer $row1
     * @param integer $col2
     * @param integer $row2
     * @return Worksheet
     */
    public function mergeCellsByColumnAndRow(int $col1, int $row1, int $col2, int $row2): Worksheet
    {
        $cell1 = Coordinate::stringFromColumnIndex($col1) . $row1;
        $cell2 = Coordinate::stringFromColumnIndex($col2) . $row2;
        return $this->mergeCells("{$cell1}:{$cell2}");
    }

    /**
     * 文字色または塗りつぶしの色を設定する
     *
     * @param string $color
     * @return Worksheet
     */
    public function setARGB(string $color): Worksheet
    {
        switch ($this->currentColor) {
            case 'color:':
                $this->setCellColorByColumnAndRow($this->lastColor['column'], $this->lastColor['row'], $color);
            break;
            case 'fillStart':
                $this->setCellFillColorByColumnAndRow($this->lastFillStartColor['column'], $this->lastFillStartColor['row'], $color);
            break;
            case 'fillEnd':
                $this->setCellFillColorByColumnAndRow($this->lastFillEndColor['column'], $this->lastFillEndColor['row'], null, $color);
            break;
        }
        return $this;
    }

    /**
     * 塗りつぶしのパターンを設定する
     *
     * @param string $style
     * @return Worksheet
     */
    public function setBorderStyle(string $style): Worksheet
    {
        $column = $this->lastBorder['column'];
        $row = $this->lastBorder['row'];
        $positon = $this->lastBorderPosition;
        return $this->setCellBorderStyleByColumnAndRow($column, $row, $positon, $style);
    }

    /**
     * セルに値を設定する
     *
     * @param string $cell
     * @param string $value
     * @return Worksheet
     */
    public function setCellValue(string $cell, $value): Worksheet
    {
        list($column, $row) = Coordinate::columnRowIndexFromString($cell);
        return $this->setCellValueByColumnAndRow($column, $row, $value);
    }

    /**
     * セルに値を設定する
     *
     * @param integer $column
     * @param integer $row
     * @param string $value
     * @return Worksheet
     */
    public function setCellValueByColumnAndRow(int $column, int $row, $value): Worksheet
    {
        $this->setDimension($column, $row);
        if (strpos($value, '=') === 0) {
            $this->data[$row][$column]['formula'] = substr($value, 1);
            unset($this->data[$row][$column]['value']);
        } else {
            if (isset($this->spreadsheet->strings[$value])) {
                $index = $this->spreadsheet->strings[$value];
            } else {
                $index = count($this->spreadsheet->strings);
                $this->spreadsheet->strings[$value] = $index;
            }
            $this->data[$row][$column]['value'] = $index;
            unset($this->data[$row][$column]['formula']);
        }
        return $this;
    }

    /**
     * 塗りつぶしのパターンを設定する
     *
     * @param string $fillType
     * @return Worksheet
     */
    public function setFillType($fillType): Worksheet
    {
        $column = $this->lastFill['column'];
        $row = $this->lastFill['row'];
        return $this->setCellFillTypeByColumnAndRow($column, $row, $fillType);
    }

    // ----------------------------------------------------------------

    /**
     * シートデータを取得する
     *
     * @return array
     */
    public function getSheetData(): array
    {
        return $this->data;
    }

    /**
     * xlsxファイルに書き込むxml文字列を返す
     *
     * @return string
     */
    public function getXml(): string
    {
        $writer = Xml::getWriter();
        $writer->startDocument('1.0', 'UTF-8', 'yes');
        Xml::start(
            'worksheet',
            [
                'xml:space' => 'preserve',
                'xmlns' => 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                'xmlns:r' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                'xmlns:xdr' => 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
                'xmlns:x14' => 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',
                'xmlns:mc' => 'http://schemas.openxmlformats.org/markup-compatibility/2006',
                'mc:Ignorable' => 'x14ac',
                'xmlns:x14ac' => 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac',
            ]
        )
            ->start('sheetPr')
                ->element('outlinePr', ['summaryBelow' => '1', 'summaryRight' => '1'])
            ->end();
        $minCell = Coordinate::stringFromColumnIndex($this->dimension['min']['column'])
            . $this->dimension['min']['row'];
        $maxCell = Coordinate::stringFromColumnIndex($this->dimension['max']['column'])
            . $this->dimension['max']['row'];
        Xml::element('dimension', ['ref' => "{$minCell}:{$maxCell}"])
            ->start('sheetViews')
                ->start('sheetView', ['tabSelected' => '1', 'workbookViewId' => '0'])
                    ->element('selection', ['activeCell' => 'A1', 'sqref' => 'A1'])
                ->end()
            ->end()
            ->element('sheetFormatPr', ['defaultRowHeight' => '18.75'])
            ->start('sheetData');

        ksort($this->data);
        foreach ($this->data as $row => $rowData) {
            ksort($rowData);

            if ($this->dimension['max']['column'] === 1) {
                Xml::start('row', ['r' => $row, 'span' => '1:1']);
            } else {
                Xml::start('row', ['r' => $row, 'spans' => '1:' . $this->dimension['max']['column']]);
            }

            foreach ($rowData as $column => $columnData) {

                // font配列の生成
                $fontKey = null;
                if (isset($columnData['font'])) {
                    $fontKey = json_encode($columnData['font']);
                    if (empty($this->spreadsheet->fonts[$fontKey])) {
                        $fontId = count($this->spreadsheet->fonts);
                        $fontData = array_merge(
                            ['id' => $fontId],
                            $columnData['font']
                        );
                        $this->spreadsheet->fonts[$fontKey] = $fontData;
                    }
                }

                // numFmt配列の生成
                $numFmtKey = null;

                // fill配列の生成
                $fillKey = null;
                if (isset($columnData['fill'])) {
                    $fillKey = json_encode($columnData['fill']);
                    if (empty($this->spreadsheet->fills[$fillKey])) {
                        $fillId = count($this->spreadsheet->fills);
                        $fillData = array_merge(
                            ['id' => $fillId],
                            $columnData['fill']
                        );
                        $this->spreadsheet->fills[$fillKey] = $fillData;
                    }
                }

                // border配列の生成
                $borderKey = null;
                if (isset($columnData['border'])) {
                    $borderKey = json_encode($columnData['border']);
                    if (empty($this->spreadsheet->borders[$borderKey])) {
                        $borderId = count($this->spreadsheet->borders);
                        $borderData = array_merge(
                            ['id' => $borderId],
                            $columnData['border']
                        );
                        $this->spreadsheet->borders[$borderKey] = $borderData;
                    }
                }
                // cellXf
                $cellXfKey = $fontKey . $numFmtKey . $fillKey . $borderKey;
                if (! empty($cellXfKey) && empty($this->spreadsheet->cellXfs[$cellXfKey])) {
                    $cellXfId = count($this->spreadsheet->cellXfs);
                    $this->spreadsheet->cellXfs[$cellXfKey] = [
                        'id'       => $cellXfId,
                        'fontId'   => isset($fontKey) ? $this->spreadsheet->fonts[$fontKey]['id'] : 0,
                        'fillId'   => isset($fillKey) ? $this->spreadsheet->fills[$fillKey]['id'] : 0,
                        'borderId' => isset($borderKey) ? $this->spreadsheet->borders[$borderKey]['id'] : 0,
                    ];
                }
                $cell = Coordinate::stringFromColumnIndex($column) . $row;
                if (empty($cellXfKey)) {
                    // 書式未設定
                    Xml::start('c', ['r' => $cell, 't' => 's']);
                } else {
                    Xml::start('c', ['r' => $cell, 's' => $this->spreadsheet->cellXfs[$cellXfKey]['id'], 't' => 's']);
                }
                isset($columnData['formula']) and Xml::element('f', null, $columnData['formula']);
                isset($columnData['value']) and Xml::element('v', null, $columnData['value']);
                Xml::end();
            }
            Xml::end();
        }
        Xml::end(); // </sheetData>

        if (count($this->mergeCells) > 0) {
            Xml::start('mergeCells');
            foreach ($this->mergeCells as $cell) {
                Xml::element('mergeCell', ['ref' => $cell]);
            }
            Xml::end();
        }
        Xml::element(
            'pageMargins',
            [
                'left' => '0.7', 'right' => '0.7',
                'top' => '0.75', 'bottom' => '0.75',
                'header' => '0.3', 'footer' => '0.3'
            ]
        )
        ->element('pageSetup');
        Xml::end(); // </worksheet>
        return $writer->outputMemory();
    }

    /**
     * 指定したセルの枠線パターンを設定する
     *
     * @param string $cell
     * @param string $position
     * @param string $style
     * @return Worksheet
     */
    public function  setCellBorderStyle(string $cell, string $position, string $style): Worksheet
    {
        list($column, $row) = Coordinate::columnRowIndexFromString($cell);
        return $this->setCellBorderStyleByColumnAndRow($column, $row, $position, $style);
    }

    /**
     * 指定したセルの枠線パターンを設定する
     *
     * @param integer $column
     * @param integer $row
     * @param string $position
     * @param string $style
     * @return Worksheet
     */
    public function setCellBorderStyleByColumnAndRow(int $column, int $row, string $position, string $style): Worksheet
    {
        $this->setDimension($column, $row);
        if (empty($this->data[$row][$column]['border'][$position]['style'])) {
            $this->data[$row][$column]['border'][$position]['style'] = $style;
        }
        return $this;
    }

    /**
     * 指定したセルに文字色をつける
     *
     * @param string $cell
     * @param string $argb
     * @return Worksheet
     */
    public function setCellColor(string $cell, string $argb): Worksheet
    {
        list($column, $row) = Coordinate::columnRowIndexFromString($cell);
        return $this->setCellColorByColumnAndRow($column, $row, $argb);
    }

    /**
     * 指定したセルに文字色をつける
     *
     * @param integer $column
     * @param integer $row
     * @param string $argb
     * @return Worksheet
     */
    public function setCellColorByColumnAndRow(int $column, int $row, string $argb): Worksheet
    {
        $this->setDimension($column, $row);
        if (empty($this->data[$row][$column]['font']['sz'])) {
            $this->data[$row][$column]['font']['sz'] = $this->spreadsheet->defaultFont['sz'];
        }
        $this->data[$row][$column]['font']['color'] = $argb;
        if (empty($this->data[$row][$column]['font']['name'])) {
            $this->data[$row][$column]['font']['name'] = $this->spreadsheet->defaultFont['name'];
        }
        return $this;
    }

    /**
     * 指定したセルを塗りつぶす
     *
     * @param string $cell
     * @param string $argb
     * @return Worksheet
     */
    public function setCellFillColor(string $cell, string $startColor, string $encColor = null): Worksheet
    {
        list($column, $row) = Coordinate::columnRowIndexFromString($cell);
        return $this->setCellFillColorByColumnAndRow($column, $row, $startColor, $encColor);
    }

    /**
     * 指定したセルを塗りつぶす
     *
     * @param integer $column
     * @param integer $row
     * @param mixed $startColor
     * @param string $endColor
     * @return Worksheet
     */
    public function setCellFillColorByColumnAndRow(int $column, int $row, $startColor, string $endColor = null): Worksheet
    {
        $this->setDimension($column, $row);
        if (isset($startColor)) {
            $this->data[$row][$column]['fill']['fgColor'] = $startColor;
        }
        if (isset($endColor)) {
            $this->data[$row][$column]['fill']['bgColor'] = $endColor;
        }
        return $this;
    }

    /**
     * 指定したセルの塗りつぶしタイプを設定する
     *
     * @param string $cell
     * @param string $fillType
     * @return Worksheet
     */
    public function setCellFillType(string $cell, string $fillType): Worksheet
    {
        list($column, $row) = Coordinate::columnRowIndexFromString($cell);
        return $this->setCellFillTypeByColumnAndRow($column, $row, $fillType);
    }

    /**
     * 指定したセルの塗りつぶしタイプを設定する
     *
     * @param integer $column
     * @param integer $row
     * @param string $fillType
     * @return Worksheet
     */
    public function setCellFillTypeByColumnAndRow(int $column, int $row, string $fillType): Worksheet
    {
        $this->setDimension($column, $row);
        $this->data[$row][$column]['fill']['patternType'] = $fillType;
        return $this;
    }

    /**
     * シートの利用範囲を更新する
     *
     * @param int $column
     * @param int $row
     * @return void
     */
    protected function setDimension(int $column, int $row)
    {
        if (empty($this->dimension)) {
            $this->dimension['min']['column'] = $column;
            $this->dimension['min']['row'] = $row;
            $this->dimension['max']['column'] = $column;
            $this->dimension['max']['row'] = $row;
        } else {
            $this->dimension['min']['column'] = min($this->dimension['min']['column'], $column);
            $this->dimension['min']['row'] = min($this->dimension['min']['row'], $row);
            $this->dimension['max']['column'] = max($this->dimension['max']['column'], $column);
            $this->dimension['max']['row'] = max($this->dimension['max']['row'], $row);
        }
    }
}
