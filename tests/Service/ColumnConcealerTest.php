<?php


namespace Excel\Templating\Service;


class ColumnConcealerTest extends \PHPUnit_Framework_TestCase 
{
    /**
     * @test
     * @dataProvider provideSheetFileName
     * @param string $filename
     * @param array $colsToHide
     * @param string $expectedOutput
     */
    public function test_execute($filename, array $colsToHide, $expectedOutput)
    {
        $basePath = __DIR__.'/../data/xmls';
        $xmls = [
            'xl/workbook.xml' => file_get_contents($basePath.'/xl/workbook.xml'),
            'xl/_rels/workbook.xml.rels' => file_get_contents($basePath.'/xl/_rels/workbook.xml.rels'),
            'xl/worksheets/sheet1.xml' => file_get_contents($basePath.'/'.$filename.'.xml'),
        ];

        $output = $this->getMock('\ZipArchive');
        $output
            ->expects($this->any())
            ->method('getFromName')
            ->with($this->isType('string'))
            ->will($this->returnCallback(function($filename) use ($xmls){
                return $xmls[$filename];
            }))
        ;
        $output
            ->expects($this->once())
            ->method('addFromString')
            ->with($this->isType('string'), $this->callback(function($xml) use ($expectedOutput){
                $this->assertContains($expectedOutput, $xml);

                return true;
            }))
        ;

        $service = new ColumnConcealer();
        $service->execute($output, ['test' => $colsToHide]);
    }

    public function provideSheetFileName()
    {
        $colsXml1 = '<cols><col min="1" max="4" width="8"/><col min="5" max="5" width="0" customWidth="1" hidden="1"/><col min="6" max="10" width="8"/><col min="11" max="11" width="0" customWidth="1" hidden="1"/><col min="12" max="19" width="9"/><col min="20" max="20" width="0" customWidth="1" hidden="1"/></cols>';
        $colsXml2 = '<cols><col min="2" max="2" width="0" customWidth="1" hidden="1"/><col min="3" max="3" width="0" customWidth="1" hidden="1"/></cols>';

        return [
            ['sheet_with_cols', [5, 11, 20], $colsXml1],
            ['sheet_with_no_cols', [2, 3], $colsXml2],
        ];
    }
}
