<?php


namespace Excel\Templating\Service;

use Excel\Templating\Service\Util\Sheet as SheetUtil;

class RowRemover implements Service
{
    /**
     * @param \ZipArchive $output
     * @param array $rowMapsToDelete  [$sheetName => $rowsToDelete]
     */
    public function execute(\ZipArchive $output, array $rowMapsToDelete = null)
    {
        $sheetNames = array_keys($rowMapsToDelete);
        $sheetXmls = SheetUtil::convertNamesToXmls($output, $sheetNames);

        foreach ($rowMapsToDelete as $sheetName => $rowsToDelete) {
            $xmlPath = $this->formatSheetXmlPath($sheetXmls[$sheetName]);
            $dom = new \DOMDocument;
            $dom->loadXML($output->getFromName($xmlPath));
            $xpath = new \DOMXPath($dom);
            $xpath->registerNamespace('s', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

            $deletedRowCount = 0;
            foreach ($xpath->query('//s:worksheet/s:sheetData/s:row') as $element) {
                /** @var \DOMElement $element */
                if (in_array($element->getAttribute('r'), $rowsToDelete)) {
                    $element->parentNode->removeChild($element);
                    $deletedRowCount++;
                } else {
                    $oldRowNumber = $element->getAttribute('r');
                    $newRowNumber = $oldRowNumber - $deletedRowCount;
                    $element->setAttribute('r', $newRowNumber);
                    if ($element->hasChildNodes()) {
                        foreach ($element->childNodes as $columnNode) {
                            /** @var \DOMElement $columnNode */
                            $columnNode->setAttribute('r', str_replace($oldRowNumber,
                                $newRowNumber, $columnNode->getAttribute('r')));
                        }
                    }
                }
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
