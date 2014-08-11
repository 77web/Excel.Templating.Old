<?php


namespace Excel\Templating\Service;

use Excel\Templating\Service\Util\Sheet as SheetUtil;

class ColumnRemover implements Service
{
    /**
     * @param \ZipArchive $output
     * @param array $colMapsToRemove
     */
    public function execute(\ZipArchive $output, array $colMapsToRemove = null)
    {
        $sheetNames = array_keys($colMapsToRemove);
        $sheetXmls = SheetUtil::convertNamesToXmls($output, $sheetNames);

        foreach ($colMapsToRemove as $sheetName => $colsToRemove) {
            $xmlPath = $this->formatSheetXmlPath($sheetXmls[$sheetName]);
            $dom = new \DOMDocument;
            $dom->loadXML($output->getFromName($xmlPath));
            $xpath = new \DOMXPath($dom);
            $xpath->registerNamespace('s', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

            // remove element "c" from each row
            $colCharsToRemove = $this->convertColNumbersToColChars($colsToRemove);
            foreach ($xpath->query('//s:worksheet/s:sheetData/s:row') as $row) {
                /** @var \DOMElement $row */
                $removedCells = [];
                foreach ($row->childNodes as $cell) {
                    if (!($cell instanceOf \DOMElement) || $cell->nodeName !== 'c') {
                        continue;
                    }
                    $colChar = preg_replace('/[0-9]+/', '', $cell->getAttribute('r'));
                    $colNumber = $this->convertColCharToColNumber($colChar);
                    if (in_array($colChar, $colCharsToRemove)) {
                        $removedCells[] = $cell;
                    } else {
                        $removedCellCount = $this->countRemovedCells($colNumber, $colsToRemove);
                        if ($removedCellCount > 0) {
                            $newColChar = $this->convertColNumberToColChar($colNumber - $removedCellCount);
                            $cell->setAttribute('r', str_replace($colChar, $newColChar, $cell->getAttribute('r')));
                        }
                    }
                }

                foreach ($removedCells as $cell) {
                    $row->removeChild($cell);
                }
            }
            // if the sheet has "cols" entry, fix it
            $elements = $xpath->query('//s:worksheet/s:cols');
            if (0 !== $elements->length) {
                /** @var \DOMElement $cols */
                $cols = $elements->item(0);
                foreach ($cols->childNodes as $col) {
                    if (!($col instanceOf \DOMElement) || $col->nodeName !== 'col') {
                        continue;
                    }
                    $min = $col->getAttribute('min');
                    $max = $col->getAttribute('max');
                    $originalMin = $min;
                    $originalMax = $max;

                    foreach ($colsToRemove as $colNumber) {
                        if ((int)$colNumber < (int)$originalMin) {
                            $min--;
                        }
                        if ((int)$colNumber <= (int)$originalMax) {
                            $max--;
                        }
                    }
                    $col->setAttribute('min', $min);
                    $col->setAttribute('max', $max);
                }
            }
            $output->addFromString($xmlPath, $dom->saveXML());
        }
    }

    /**
     * @param string $xmlFileName
     * @return string
     */
    private function formatSheetXmlPath($xmlFileName)
    {
        return sprintf('xl/worksheets/%s', $xmlFileName);
    }

    /**
     * @param array $colNumbers
     * @return array
     */
    private function convertColNumbersToColChars(array $colNumbers)
    {
        $chars = [];
        foreach ($colNumbers as $colNumber) {
            $chars[] = $this->convertColNumberToColChar($colNumber);
        }

        return $chars;
    }

    /**
     * @param int $colNumber
     * @return string
     */
    private function convertColNumberToColChar($colNumber)
    {
        $charA = 65;

        $char = '';

        $quot = intval($colNumber / 26);
        $mod = $colNumber % 26;
        if ($quot > 0) {
            $char .= chr($charA + $quot - 1);
        }
        if ($mod > 0) {
            $char .= chr($charA + $mod - 1);
        }

        return $char;
    }

    /**
     * @param string $colChar
     * @return int
     */
    private function convertColCharToColNumber($colChar)
    {
        $charA = 65;

        if (strlen($colChar) == 1) {
            return ord($colChar) - $charA + 1;
        }

        $quot = ord($colChar[0]) - $charA + 1;
        $mod = ord($colChar[1]) - $charA + 1;

        return $quot * 26 + $mod;
    }

    /**
     * @param int $colNumber
     * @param array $colNumbersToRemove
     * @return int
     */
    private function countRemovedCells($colNumber, array $colNumbersToRemove)
    {
        $count = 0;
        foreach ($colNumbersToRemove as $colNumberToRemove) {
            if ($colNumberToRemove < $colNumber) {
                $count++;
            }
        }

        return $count;
    }
}
