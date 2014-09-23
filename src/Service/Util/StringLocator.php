<?php

namespace Excel\Templating\Service\Util;

use Excel\Templating\Service\Util\Sheet as SheetUtil;

/**
 * Class StringLocator
 */
class StringLocator
{
    /**
     * @var SharedStringReader
     */
    private $reader;

    /**
     * @param SharedStringReader $reader
     */
    public function __construct(SharedStringReader $reader)
    {
        $this->reader = $reader;
    }

    /**
     * @param \ZipArchive $zip
     * @param string $sheetName
     * @return array
     */
    public function getLocation(\ZipArchive $zip, $sheetName)
    {
        $sharedStrings = $this->reader->read($zip);

        $sheets = SheetUtil::convertNamesToXmls($zip, [$sheetName]);
        $xmlPath = $sheets[$sheetName];

        return $this->makeMap($zip->getFromName($xmlPath), $sharedStrings);
    }

    /**
     * @param string $xml
     * @param array $sharedStrings
     * @return array
     */
    private function makeMap($xml, array $sharedStrings)
    {
        $dom = new \DOMDocument;
        $dom->loadXML($xml);
        $xpath = new \DOMXPath($dom);
        $xpath->registerNamespace('s', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

        $map = [];
        foreach ($xpath->query('//s:worksheet/s:sheetData/s:row/s:c[@t="s"]/s:v') as $stringCellValue) {
            $stringId = $stringCellValue->nodeValue;
            if (!isset($sharedStrings[$stringId])) {
                continue;
            }

            $map[$stringCellValue->parentNode->getAttribute('r')] = $sharedStrings[$stringId];
        }

        return $map;
    }
} 
