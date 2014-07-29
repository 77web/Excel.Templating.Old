<?php


namespace Excel\Templating\Service\Util;


class SheetTest extends \PHPUnit_Framework_TestCase 
{
    /**
     * @param array $names
     * @param array $expectedRelIds
     * @dataProvider provideNameToRelIdData
     */
    public function test_convertNamesToRelIds($names, $expectedRelIds)
    {
        $dummyXml = file_get_contents(__DIR__.'/../../data/xmls/xl/workbook.xml');

        $zip = $this->getMock('\ZipArchive');
        $zip
            ->expects($this->once())
            ->method('getFromName')
            ->with('xl/workbook.xml')
            ->will($this->returnValue($dummyXml))
        ;

        $relIds = Sheet::convertNamesToRelIds($zip, $names);
        $this->assertEquals($expectedRelIds, array_values($relIds));
    }

    /**
     * @return array
     */
    public function provideNameToRelIdData()
    {
        return [
            [['test'], ['rId1']],
            [['Sheet2'], ['rId2']],
            [['test', 'Sheet2'], ['rId1', 'rId2']],
        ];
    }

    /**
     * @param array $relIds
     * @param array $expectedXmls
     * @dataProvider provideRelIdToXmlData
     */
    public function test_convertRelIdsToXmls($relIds, $expectedXmls)
    {
        $dummyRelXml = file_get_contents(__DIR__.'/../../data/xmls/xl/_rels/workbook.xml.rels');

        $zip = $this->getMock('\ZipArchive');
        $zip
          ->expects($this->once())
          ->method('getFromName')
          ->with('xl/_rels/workbook.xml.rels')
          ->will($this->returnValue($dummyRelXml))
        ;

        $xmls = Sheet::convertRelIdsToXmls($zip, $relIds);
        $this->assertEquals($expectedXmls, array_values($xmls));
    }

    /**
     * @return array
     */
    public function provideRelIdToXmlData()
    {
        return [
          [['rId1'], ['sheet1.xml']],
          [['rId2'], ['sheet2.xml']],
          [['rId1', 'rId2'], ['sheet1.xml', 'sheet2.xml']],
        ];
    }

    /**
     * @param array $names
     * @param array $expectedXmls
     * @dataProvider provideNameToXmlData
     */
    public function test_convertNamesToXmls($names, $expectedXmls)
    {
        $dummyXml = file_get_contents(__DIR__.'/../../data/xmls/xl/workbook.xml');
        $dummyRelXml = file_get_contents(__DIR__.'/../../data/xmls/xl/_rels/workbook.xml.rels');

        $zip = $this->getMock('\ZipArchive');
        $zip
          ->expects($this->exactly(2))
          ->method('getFromName')
          ->with($this->logicalOr('xl/workbook.xml', 'xl/_rels/workbook.xml.rels'))
          ->will($this->onConsecutiveCalls($dummyXml, $dummyRelXml))
        ;

        $relIds = Sheet::convertNamesToXmls($zip, $names);
        $this->assertEquals($expectedXmls, array_values($relIds));
    }

    /**
     * @return array
     */
    public function provideNameToXmlData()
    {
        return [
          [['test'], ['sheet1.xml']],
          [['Sheet2'], ['sheet2.xml']],
          [['test', 'Sheet2'], ['sheet1.xml', 'sheet2.xml']],
        ];
    }
}
