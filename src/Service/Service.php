<?php


namespace Excel\Templating\Service;

/**
 * Interface Service
 * サービスのinterface
 *
 * @package Excel\Templating\Service
 */
interface Service
{
    public function execute(\ZipArchive $output);
} 
