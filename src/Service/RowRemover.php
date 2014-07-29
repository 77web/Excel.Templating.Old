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

            $rowNumber = 1;
            foreach ($xpath->query('//s:worksheet/s:sheetData/s:row') as $element) {
                /** @var \DOMElement $element */
                if (in_array($element->getAttribute('r'), $rowsToDelete)) {
                    $element->parentNode->removeChild($element);
                } else {
                    $oldRowNumber = $element->getAttribute('r');
                    $element->setAttribute('r', $rowNumber);
                    if ($element->hasChildNodes()) {
                        foreach ($element->childNodes as $columnNode) {
                            /** @var \DOMElement $columnNode */
                            $columnNode->setAttribute('r', str_replace($oldRowNumber,
                                $rowNumber, $columnNode->getAttribute('r')));
                        }
                    }

                    $rowNumber++;
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
