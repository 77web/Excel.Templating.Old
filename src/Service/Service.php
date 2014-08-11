<?php


namespace Excel\Templating\Service;

/**
 * Interface Service
 * interface of service
 *
 * @package Excel\Templating\Service
 */
interface Service
{
    public function execute(\ZipArchive $output);
} 
