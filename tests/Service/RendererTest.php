<?php


namespace Excel\Templating\Service;

class RendererTest extends \PHPUnit_Framework_TestCase 
{
    public function test_execute()
    {
        $dummyXml = '<a>%var%</a><b>notvar%</b><c>%notvar</c><d>normal string</d><e>%non_existant_var%</e>';
        $zip = $this->getMock('\ZipArchive');
        $zip
            ->expects($this->once())
            ->method('getFromName')
            ->with($this->isType('string'))
            ->will($this->returnValue($dummyXml))
        ;
        $zip
            ->expects($this->once())
            ->method('addFromString')
            ->with($this->isType('string'), $this->callback(function($xml){
                  $this->assertEquals('<a>who</a><b>notvar%</b><c>%notvar</c><d>normal string</d><e></e>', $xml);

                  return true;
              }))
        ;
        $variables = [
            '%var%' => 'who',
            '%rat%' => 'cat',
        ];

        $service = new Renderer();
        $service->execute($zip, $variables);
    }
}
