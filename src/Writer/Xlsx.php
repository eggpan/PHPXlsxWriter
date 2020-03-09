<?php

namespace Eggpan\PHPXlsxWriter\Writer;

use ZipArchive;
use Eggpan\PHPXlsxWriter\Exceptions\UnknownErrorException;
use Eggpan\PHPXlsxWriter\Writer\Xml;

class Xlsx
{
    protected $spreadsheet;

    public function __construct($spreadsheet)
    {
        $this->spreadsheet = $spreadsheet;
    }

    public function save($fileName)
    {
        $zip = new ZipArchive();
 
        $result = $zip->open($fileName, ZipArchive::CREATE);
        if ($result !== true) {
            throw new UnknownErrorException(var_export($result, true));
        }

        $zip->addFromString('[Content_Types].xml', $this->getContentTypeXml());
        $zip->addFromString('_rels/.rels', $this->getRels());
        $zip->addFromString('docProps/app.xml', $this->getDocPropsApp());
        $zip->addFromString('docProps/core.xml', $this->getDocPropsCore());

        $zip->addFromString('xl/sharedStrings.xml', $this->getXlSharedStrings());
        $zip->addFromString('xl/theme/theme1.xml', $this->getXlTheme());
        $zip->addFromString('xl/workbook.xml', $this->getXlWorkbook());
        $zip->addFromString('xl/_rels/workbook.xml.rels', $this->getXlWorkbookXmlRels());

        foreach ($this->spreadsheet->sheets as $index => $sheet) {
            $xml = $sheet->getXml();
            $zip->addFromString('xl/worksheets/sheet' . ($index + 1) . '.xml', $xml);
        }
        $zip->addFromString('xl/styles.xml', $this->getXlStyle());

        $zip->close();
    }

    /**
     * [Content_Types].xml
     *
     * @return string
     */
    protected function getContentTypeXml()
    {
        Xml::getWriter()->startDocument('1.0', 'UTF-8', 'yes');
        Xml::start('Types', ['xmlns' => 'http://schemas.openxmlformats.org/package/2006/content-types'])
        ->element(
            'Default',
            [
                'Extension' => 'bin',
                'ContentType' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings',
            ]
        )
        ->element(
            'Default',
            [
                'Extension' => 'rels',
                'ContentType' => 'application/vnd.openxmlformats-package.relationships+xml',
            ]
        )
        ->element(
            'Default',
            [
                'Extension' => 'xml',
                'ContentType' => 'application/xml',
            ]
        )
        ->element(
            'Override',
            [
                'PartName' => '/xl/workbook.xml',
                'ContentType' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
            ]
        );
        foreach ($this->spreadsheet->sheets as $index => $sheet) {
            Xml::element(
                'Override',
                [
                    'PartName' => '/xl/worksheets/sheet' . ($index + 1) . '.xml',
                    'ContentType' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
                ]
            );
        }
        Xml::element(
            'Override',
            [
                'PartName' => '/xl/theme/theme1.xml',
                'ContentType' => 'application/vnd.openxmlformats-officedocument.theme+xml',
            ]
        )
        ->element(
            'Override',
            [
                'PartName' => '/xl/styles.xml',
                'ContentType' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
            ]
        )
        ->element(
            'Override',
            [
                'PartName' => '/xl/sharedStrings.xml',
                'ContentType' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
            ]
        )
        ->element(
            'Override',
            [
                'PartName' => '/docProps/core.xml',
                'ContentType' => 'application/vnd.openxmlformats-package.core-properties+xml',
            ]
        )
        ->element(
            'Override',
            [
                'PartName' => '/docProps/app.xml',
                'ContentType' => 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
            ]
        )
        ->end(); // </Types>

        return Xml::getWriter()->outputMemory();
    }

    /**
     * docProps/app.xml
     *
     * @return string
     */
    protected function getDocPropsApp()
    {
        Xml::getWriter()->startDocument('1.0', 'UTF-8', 'yes');
        Xml::start(
            'Properties',
            [
                'xmlns' => 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
                'xmlns:vt' => 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
            ]
        )
        ->element('Application', null, 'Microsoft Excel')
        ->element('DocSecurity', null, '0')
        ->element('ScaleCrop', null, 'false')
        ->start('HeadingPairs')
            ->start('vt:vector', ['size' => '2', 'baseType' => 'variant'])
                ->start('vt:variant')
                    ->element('vt:lpstr', null, 'Worksheets')
                ->end()
                ->start('vt:variant')
                    ->element('vt:i4', null, '1')
                ->end()
            ->end()
        ->end()
        ->start('TitlesOfParts')
            ->start('vt:vector', ['size' => '1', 'baseType' => 'lpstr']);

        foreach ($this->spreadsheet->sheets as $index => $sheet) {
            Xml::element('vt:lpstr', null, $sheet->getName());
        }

           Xml::end()
        ->end() // </TitlesOfParts>
        ->element('Company')
        ->element('LinksUpToDate', null, 'false')
        ->element('SharedDoc', null, 'false')
        ->element('HyperlinksChanged', null, 'false')
        ->element('AppVersion', null, '12.0000')
        ->end(); // </Properties>
        return Xml::getWriter()->outputMemory();
    }

    protected function getDocPropsCore()
    {
        Xml::getWriter()->startDocument('1.0', 'UTF-8', 'yes');
        Xml::start(
            'cp:coreProperties',
            [
                'xmlns:cp' => 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
                'xmlns:dc' => 'http://purl.org/dc/elements/1.1/',
                'xmlns:dcterms' => 'http://purl.org/dc/terms/',
                'xmlns:dcmitype' => 'http://purl.org/dc/dcmitype/',
                'xmlns:xsi' => 'http://www.w3.org/2001/XMLSchema-instance',
            ]
        )
            ->element('dc:creator', null, 'Unknown Creator')
            ->element('cp:lastModifiedBy', null, 'Unknown Creator')
            ->element('dcterms:created', ['xsi:type' => 'dcterms:W3CDTF'], date('c'))
            ->element('dcterms:modified', ['xsi:type' => 'dcterms:W3CDTF'], date('c'))
        ->end();
        
        return Xml::getWriter()->outputMemory();
    }

    protected function getRels()
    {
        Xml::getWriter()->startDocument('1.0', 'UTF-8', 'yes');
        Xml::start('Relationships', ['xmlns' => 'http://schemas.openxmlformats.org/package/2006/relationships'])
        ->element(
            'Relationship',
            [
                'Id'     => 'rId3',
                'Type'   => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
                'Target' => 'docProps/app.xml',
            ]
        )
        ->element(
            'Relationship',
            [
                'Id'     => 'rId2',
                'Type'   => 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
                'Target' => 'docProps/core.xml',
            ]
        )
        ->element(
            'Relationship',
            [
                'Id'     => 'rId1',
                'Type'   => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
                'Target' => 'xl/workbook.xml',
            ]
        )
        ;
        Xml::end();

        return Xml::getWriter()->outputMemory();
    }

    protected function getXlSharedStrings()
    {
        Xml::getWriter()->startDocument('1.0', 'UTF-8', 'yes');
        Xml::start(
            'sst',
            [
                'xmlns' => 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                'uniqueCount' => count($this->spreadsheet->strings)
            ]
        );
        foreach ($this->spreadsheet->strings as $value => $index) {
            Xml::start('si')
                ->element('t', null, $value)
            ->end();
        }
        Xml::end();
        return Xml::getWriter()->outputMemory();
    }

    protected function getXlStyle()
    {
        Xml::getWriter()->startDocument('1.0', 'UTF-8', 'yes');
        Xml::start(
            'styleSheet',
            [
                'xml:space' => 'preserve',
                'xmlns' => 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            ]
        )
        ->start('fonts');
        foreach ($this->spreadsheet->fonts as $hash => $font) {
            Xml::start('font')
                ->element('sz', ['val' => $font['sz']])
                ->element('color', ['rgb' => $font['color']])
                ->element('name', ['val' => $font['name']]);
            Xml::end();
        }
        Xml::end(); 

        // fills
        Xml::start('fills');
        foreach ($this->spreadsheet->fills as $hash => $fill) {
            Xml::start('fill');
                Xml::start('patternFill', ['patternType' => $fill['patternType']]);
                if (isset($fill['fgColor'])) {
                    Xml::element('fgColor', ['rgb' => $fill['fgColor']]);
                }
                if (isset($fill['bgColor'])) {
                    Xml::element('bgColor', ['rgb' => $fill['bgColor']]);
                }
                Xml::end();
            Xml::end();
        }
        Xml::end(); 

        // borders
        Xml::start('borders');
        foreach ($this->spreadsheet->borders as $hash => $border) {
            Xml::start('border');
                if (isset($border['left'])) {
                    Xml::start('left', ['style' => $border['left']['style']]);
                        Xml::element(
                            'color',
                            [
                                'rgb' => $border['left']['color'] ?? 'FF000000'
                            ]
                        );
                    Xml::end();
                }
                if (isset($border['right'])) {
                    Xml::start('right', ['style' => $border['right']['style']]);
                    Xml::element(
                        'color',
                        [
                            'rgb' => $border['right']['color'] ?? 'FF000000'
                        ]
                    );
                    Xml::end();
                }
                if (isset($border['top'])) {
                    Xml::start('top', ['style' => $border['top']['style']]);
                    Xml::element(
                        'color',
                        [
                            'rgb' => $border['top']['color'] ?? 'FF000000'
                        ]
                    );
                    Xml::end();
                }
                if (isset($border['bottom'])) {
                    Xml::start('bottom', ['style' => $border['bottom']['style']]);
                    Xml::element(
                        'color',
                        [
                            'rgb' => $border['bottom']['color'] ?? 'FF000000'
                        ]
                    );
                    Xml::end();
                }
            Xml::end();
        }
        Xml::end(); 

        Xml::start('cellStyleXfs', ['count' => '1'])
            ->element(
                'xf',
                [
                    'numFmtId' => '0',
                    'fontId' => '0',
                    'fillId' => '0',
                    'borderId' => '0',
                ]
            )
        ->end();

        Xml::start('cellXfs', ['count' => count($this->spreadsheet->cellXfs)]);
        foreach ($this->spreadsheet->cellXfs as $cellXf) {
            Xml::start(
                'xf',
                [
                    'xfId' => '0',
                    'fontId' => $cellXf['fontId'],
                    'numFmtId' => '0',
                    'fillId' => $cellXf['fillId'],
                    'borderId' => $cellXf['borderId'],
                    'applyFont' => '0',
                    'applyNumberFormat' => '0',
                    'applyFill' => '0',
                    'applyBorder' => '0',
                    'applyAlignment' => '0',
                    
                ]
            )
                ->element(
                    'alignment',
                    [
                        'horizontal' => 'general',
                        'vertical' => 'bottom',
                        'textRotation' => '0',
                        'wrapText' => 'false',
                        'shrinkToFit' => 'false',
                    ]
                )
            ->end();
        }
        Xml::end();

        Xml::start('cellStyles', ['count' => 1])
            ->element('cellStyle', ['name' => 'Normal', 'xfId' => '0', 'builtinId' => '0'])
        ->end()
        ->element('dxfs', ['count' => 0])
        ->element(
            'tableStyles',
            [
                'defaultTableStyle' => 'TableStyleMedium9',
                'defaultPivotStyle' => 'PivotTableStyle1',
            ]
        );

        Xml::end(); // </styleSheet>
        return Xml::getWriter()->outputMemory();
    }

    protected function getXlTheme()
    {
        Xml::getWriter()->startDocument('1.0', 'UTF-8', 'yes');
        Xml::getWriter()->writeRaw('<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>');
        return Xml::getWriter()->outputMemory();
    }

    protected function getXlWorkbook()
    {
        Xml::getWriter()->startDocument('1.0', 'UTF-8', 'yes');
        Xml::start(
            'workbook',
            [
                'xml:space' => 'preserve',
                'xmlns' => 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                'xmlns:r' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            ]
        )
        ->element(
            'fileVersion',
            [
                'appName' =>  'xl',
                'lastEdited' => '4',
                'lowestEdited' => '4',
                'rupBuild' => '4505',
            ]
        )
        ->element(
            'workbookPr',
            [
                'codeName' =>  'ThisWorkbook',
            ]
        )
        ->start('bookViews')
            ->element('workbookView')
        ->end()
        ->start('sheets');
        foreach ($this->spreadsheet->sheets as $index => $sheet) {
            Xml::element(
                'sheet',
                [
                    'name' => $sheet->getName(),
                    'sheetId' => ($index + 1),
                    'r:id' => 'rId' . ($index + 4),
                ]
            );
        }
        
        Xml::end() // </sheets>
        ->element('calcPr', ['calcId' => '999999'])
        ->end(); // </workbook>
        return Xml::getWriter()->outputMemory();
    }

    protected function getXlWorkbookXmlRels()
    {
        Xml::getWriter()->startDocument('1.0', 'UTF-8', 'yes');
        Xml::start('Relationships', ['xmlns' => 'http://schemas.openxmlformats.org/package/2006/relationships'])
        ->element(
            'Relationship',
            [
                'Id'     => 'rId1',
                'Type'   => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
                'Target' => 'styles.xml',
            ]
        )
        ->element(
            'Relationship',
            [
                'Id'     => 'rId2',
                'Type'   => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
                'Target' => 'theme/theme1.xml',
            ]
        )
        ->element(
            'Relationship',
            [
                'Id'     => 'rId3',
                'Type'   => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
                'Target' => 'sharedStrings.xml',
            ]
        );
        foreach ($this->spreadsheet->sheets as $index => $sheet) {
            Xml::element(
                'Relationship',
                [
                    'Id' => 'rId' . ($index + 4),
                    'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                    'Target' => 'worksheets/sheet' . ($index + 1) . '.xml',
                ]
            );
        }
        Xml::end();
        return Xml::getWriter()->outputMemory();
    }

}
