<?php

namespace Eggpan\PHPXlsxWriter;

use Eggpan\PHPXlsxWriter\Exceptions\SheetNotFoundException;
use Eggpan\PHPXlsxWriter\Worksheet\Worksheet;
use InvalidArgumentException;

class Spreadsheet
{
    public $sheets = [];
    public $strings = [];
    public $stringCount = 0;
    public $defaultFont = [
        'sz'    => '11',
        'color' => 'FF000000',
        'name'  => 'Calibri'
    ];
    public $fonts = [
        '{"sz":"11","color":"FF000000","name":"Calibri"}' => [
            'id'    => 0,
            'sz'    => '11',
            'color' => 'FF000000',
            'name'  => 'Calibri'
        ],
    ];
    public $fills = [
        '{"id":0,"patternType":"none"}' => [
            'id'          => 0,
            'patternType' => 'none',
        ],
        '{"id":1,"patternType":"gray125"}' => [
            'id'          => 1,
            'patternType' => 'gray125',
        ]
    ];

    public $borders = [
        '{"id":0}' => [
            'id'    => 0,
        ],
    ];

    public $cellXfs = [
        [
            'fontId' => '0',
            'fillId' => '0',
            'borderId' => '0',
        ]
    ];
    protected $activeSheetIndex;

    public function __construct()
    {
        $this->createSheet();
    }

    public function addSheet(Worksheet $worksheet, $index = null)
    {
        if (isset($index)) {
            if (empty($this->sheets[$index]) && $index > count($this->sheets)) {
                throw new InvalidArgumentException("index値が正しくありません。");
            }
            if ($index === count($this->sheets)) {
                $this->sheets[$index] = $worksheet;
            } else {
                array_splice($this->sheets, $index, 0, [$worksheet]);
            }
        } else {
            $index = count($this->sheets);
            $this->sheets[$index] = $worksheet;
        }
        $this->activeSheetIndex = $index;
    }

    public function createSheet()
    {
        $index = count($this->sheets);
        $sheetName = 'Sheet' . ($index + 1);
        $worksheet = new Worksheet($this, $sheetName);
        $this->addSheet($worksheet, $index);
    }

    public function getActiveSheet(): Worksheet
    {
        return $this->sheets[$this->activeSheetIndex];
    }

    public function getSheet($index): Worksheet
    {
        if (empty($this->sheet[$index])) {
            throw new SheetNotFoundException();
        }
        return $this->sheet[$index];
    }

    //public function getSheetByName($name)

    public function setActiveSheetIndex($index)
    {
        if (empty($this->sheet[$index])) {
            throw new SheetNotFoundException();
        }
        $this->activeSheetIndex = $index;
    }
}
