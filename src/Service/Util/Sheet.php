<?php


namespace Excel\Templating\Service\Util;


class Sheet
{
    /**
     * @param \ZipArchive $excel
     * @param array $names
     * @return array
     */
    public static function convertNamesToRelIds(\ZipArchive $excel, array $names)
    {
        $dom = new \DOMDocument;
        $dom->loadXML($excel->getFromName('xl/workbook.xml'));
        $xpath = new \DOMXPath($dom);
        $xpath->registerNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        $xpath->registerNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationshipts');
        $relIds = [];
        foreach ($names as $name) {
            $elements = $xpath->query('//x:workbook/x:sheets/x:sheet[@name="'.$name.'"]');
            if (1 === $elements->length) {
                /** @var \DOMElement $element */
                $element = $elements->item(0);
                $relIds[] = $element->getAttribute('r:id');
            }
        }

        $dom = null;
        $xpath = null;

        return $relIds;
    }

    /**
     * @param \ZipArchive $excel
     * @param array $names
     * @return array
     */
    public static function convertNamesToXmls(\ZipArchive $excel, array $names)
    {
        $relIds = self::convertNamesToRelIds($excel, $names);

        return self::convertRelIdsToXmls($excel, $relIds);
    }

    /**
     * @param \ZipArchive $excel
     * @param array $relIds
     * @return array
     */
    public static function convertRelIdsToXmls(\ZipArchive $excel, array $relIds)
    {
        $dom = new \DOMDocument;
        $dom->loadXML($excel->getFromName('xl/_rels/workbook.xml.rels'));
        $xpath = new \DOMXPath($dom);
        $xpath->registerNamespace('r', 'http://schemas.openxmlformats.org/package/2006/relationships');
        $xmls = [];
        foreach ($relIds as $relId) {
            $elements = $xpath->query('//r:Relationships/r:Relationship[@Id="'.$relId.'"]');
            if (1 === $elements->length) {
                /** @var \DOMElement $element */
                $element = $elements->item(0);
                $xmls[] = basename($element->getAttribute('Target'));
            }
        }

        $dom = null;
        $xpath = null;

        return $xmls;
    }
}
