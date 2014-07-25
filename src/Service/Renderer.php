<?php


namespace Excel\Templating\Service;

/**
 * Class Renderer
 * テンプレート上の変数を値に置き換える
 *
 * @package Excel\Templating\Service
 */
class Renderer implements Service
{
    private static $targetXmlPath = 'xl/sharedStrings.xml';

    public function execute(\ZipArchive $output, array $variables = null)
    {
        $xml = $output->getFromName(self::$targetXmlPath);

        // 変数を置き換える
        $xml = strtr($xml, $variables);

        // $variablesに含まれない変数を空白化
        $xml = preg_replace("/%[^%<]+%/", '', $xml);

        $output->addFromString(self::$targetXmlPath, $xml);
    }
}
