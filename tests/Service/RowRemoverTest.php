<?php


namespace Excel\Templating\Service;


class RowRemoverTest extends \PHPUnit_Framework_TestCase 
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
          'xl/worksheets/sheet1.xml' => file_get_contents($basePath.'/xl/worksheets/sheet1.xml'),
        ];
    }

    public function test_execute()
    {
        $xmls = $this->xmls;
        $zip = $this->getMock('\ZipArchive');
        $zip
          ->expects($this->any())
          ->method('getFromName')
          ->will($this->returnCallback(function($arg) use ($xmls){
                  return $xmls[$arg];
              }))
        ;
        $zip
          ->expects($this->exactly(1))
          ->method('addFromString')
          ->with($this->isType('string'), $this->isType('string'))
        ;

        $rowMaps = [
            'test' => [1, 2]
        ];

        $service = new RowRemover();
        $service->execute($zip, $rowMaps);
    }
}
