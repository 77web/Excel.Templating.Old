<?php


namespace Excel\Templating\Service;

use Excel\Templating\Service\Util\Sheet as SheetUtil;

class ColumnConcealer implements Service
{
    /**
     * @param \ZipArchive $output
     * @param array $colMapsToHide
     */
    public function execute(\ZipArchive $output, array $colMapsToHide = null)
    {
        $sheetNames = array_keys($colMapsToHide);
        $sheetXmls = SheetUtil::convertNamesToXmls($output, $sheetNames);

        foreach ($colMapsToHide as $sheetName => $colsToHide) {
            $xmlPath = $this->formatSheetXmlPath($sheetXmls[$sheetName]);
            $dom = new \DOMDocument;
            $dom->loadXML($output->getFromName($xmlPath));
            $xpath = new \DOMXPath($dom);
            $xpath->registerNamespace('s', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

            $elements = $xpath->query('//s:worksheet/s:cols');
            if (1 === $elements->length) {

                /** @var \DOMElement $cols */
                $cols = $elements->item(0);

                foreach ($colsToHide as $colNumber) {
                    $inRange = false;

                    foreach ($cols->childNodes as $col) {
                       if (!$col instanceOf \DOMElement) {
                           continue;
                       }
                       $min = $col->getAttribute('min');
                       $max = $col->getAttribute('max');
                       if ($colNumber >= $min && $colNumber <= $max) {
                           $inRange = true;

                           if ($colNumber != $min) {
                               $colToPrepend = $this->createClonedColumnElement($dom, $min, $colNumber - 1, $col);
                               $cols->insertBefore($colToPrepend, $col);
                           }

                           $hiddenCol = $this->createHiddenColumnElement($dom, $colNumber);
                           if ($colNumber == $max) {
                               $cols->replaceChild($hiddenCol, $col);
                           } else {
                               $cols->insertBefore($hiddenCol, $col);
                               $col->setAttribute('min', $colNumber + 1);
                           }
                       }
                    }

                    if (!$inRange) {
                        $col = $this->createHiddenColumnElement($dom, $colNumber);
                        $cols->appendChild($col);
                    }
                }
            } else {
                $cols = $dom->createElement('cols');

                foreach ($colsToHide as $columnNumber) {
                    $col = $this->createHiddenColumnElement($dom, $columnNumber);
                    $cols->appendChild($col);
                }

                $worksheet = $xpath->query('//s:worksheet')->item(0);
                $worksheet->appendChild($cols);
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

    /**
     * @param \DOMDocument $dom
     * @param int $columnNumber
     * @return \DOMElement
     */
    private function createHiddenColumnElement(\DOMDocument $dom, $columnNumber)
    {
        $col = $dom->createElement('col');
        $col->setAttribute('min', $columnNumber);
        $col->setAttribute('max', $columnNumber);
        $col->setAttribute('width', 0);
        $col->setAttribute('customWidth', 1);
        $col->setAttribute('hidden', 1);

        return $col;
    }

    /**
     * @param \DOMDocument $dom
     * @param int $min
     * @param int $max
     * @param \DOMElement $baseCol
     * @return \DOMElement
     */
    private function createClonedColumnElement(\DOMDocument $dom, $min, $max, \DOMElement $baseCol)
    {
        /** @var \DOMElement $col */
        $col = $baseCol->cloneNode(true);
        $col->setAttribute('min', $min);
        $col->setAttribute('max', $max);

        return $col;
    }
} 
