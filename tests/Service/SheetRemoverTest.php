<?php


namespace Excel\Templating\Service;


class SheetRemoverTest extends \PHPUnit_Framework_TestCase 
{
    /**
     * @var array
     */
    private $xmls;

    public function setUp()
    {
        $basePath = __DIR__.'/../data/xmls';
        $this->xmls = [
            'xl/workbook.xml' => file_get_contents($basePath.'/xl/workbook.xml'),
            'xl/_rels/workbook.xml.rels' => file_get_contents($basePath.'/xl/_rels/workbook.xml.rels'),
            '[Content_Types].xml' => file_get_contents($basePath.'/[Content_Types].xml'),
        ];
    }

    public function test_execute()
    {
        $xmls = $this->xmls;
        $zip = $this->getMock('\ZipArchive');
        $zip
            ->expects($this->exactly(3))
            ->method('getFromName')
            ->will($this->returnCallback(function($arg) use ($xmls){
                return $xmls[$arg];
            }))
        ;
        $zip
            ->expects($this->exactly(3))
            ->method('addFromString')
            ->with($this->isType('string'), $this->isType('string'))
        ;
        $zip
            ->expects($this->exactly(2))
            ->method('deleteName')
            ->with($this->callback(function($path){
                $this->assertContains('sheet1.xml', $path);

                return true;
            }))
        ;
        $sheetNames = [
            'test'
        ];
        
        $service = new SheetRemover();
        $service->execute($zip, $sheetNames);
    }
}
