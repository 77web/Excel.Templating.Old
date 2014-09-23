<?php

namespace Excel\Templating\Service\Util;

/**
 * Class SharedStringReader
 */
class SharedStringReader
{
    public function read(\ZipArchive $zip)
    {
        $xml = $zip->getFromName('xl/sharedStrings.xml');

        $dom = new \DOMDocument;
        $dom->loadXML($xml);
        $xpath = new \DOMXPath($dom);
        $xpath->registerNamespace('s', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

        $strings = [];
        $count = 0;
        foreach ($xpath->query('//s:sst/s:si/s:t') as $t) {
            /** @var \DOMElement $t */
            $strings[$count] = $t->nodeValue;
            $count++;
        }

        return $strings;
    }
}
