<?php


namespace Excel\Templating\Service;


class ColumnRemoverTest extends \PHPUnit_Framework_TestCase 
{
    /**
     * @test
     * @dataProvider provideTestData
     * @param string $filename
     * @param array $colsToRemove
     * @param string $expectedOutput
     */
    public function test_execute($filename, array $colsToRemove, $expectedOutput)
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

        $service = new ColumnRemover();
        $service->execute($output, ['test' => $colsToRemove]);
    }

    public function provideTestData()
    {
        $xml1 = '<cols><col min="1" max="2" width="8"/><col min="3" max="3" width="9"/></cols>';
        $xml2 = '<c r="A1" t="s"><v>1</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>1</v></c><c r="D1" t="s"><v>1</v></c><c r="E1" t="s"><v>1</v></c></row>';

        return [
            ['col_remover', [2, 4, 6], $xml1],
            ['col_remover', [1], $xml2],
        ];
    }
}
