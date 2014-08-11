<?php


namespace Excel\Templating\Service;

/**
 * Class Renderer
 * replace values
 *
 * @package Excel\Templating\Service
 */
class Renderer implements Service
{
    private static $targetXmlPath = 'xl/sharedStrings.xml';

    public function execute(\ZipArchive $output, array $variables = null)
    {
        $xml = $output->getFromName(self::$targetXmlPath);

        // replace given variables
        $xml = strtr($xml, $variables);

        // replace not given variables into empty string
        $xml = preg_replace("/%[^%<]+%/", '', $xml);

        $output->addFromString(self::$targetXmlPath, $xml);
    }
}
