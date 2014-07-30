<?php

namespace Excel\Templating\Service;

use Excel\Templating\Service\Util\Sheet as SheetUtil;

class RowConcealer implements Service
{
    /**
     * @param \ZipArchive $output
     * @param array $rowMapsToHide  [$sheetName => $rowsToHide]
     */
    public function execute(\ZipArchive $output, array $rowMapsToHide = null)
    {
        $sheetNames = array_keys($rowMapsToHide);
        $sheetXmls = SheetUtil::convertNamesToXmls($output, $sheetNames);

        foreach ($rowMapsToHide as $sheetName => $rowsToHide) {
            $xmlPath = $this->formatSheetXmlPath($sheetXmls[$sheetName]);
            $dom = new \DOMDocument;
            $dom->loadXML($output->getFromName($xmlPath));
            $xpath = new \DOMXPath($dom);
            $xpath->registerNamespace('s', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

            $rowNumber = 1;
            foreach ($xpath->query('//s:worksheet/s:sheetData/s:row') as $element) {
                if (in_array($rowNumber, $rowsToHide)) {
                    /** @var \DOMElement $element */
                    $element->setAttribute('hidden', 1);
                }

                $rowNumber++;
            }

            $output->addFromString($xmlPath, $dom->saveXML());

            $dom = null;
            $xpath = null;
        }
    }

    private function formatSheetXmlPath($xmlFileName)
    {
        return sprintf('xl/worksheets/%s', $xmlFileName);
    }
}
