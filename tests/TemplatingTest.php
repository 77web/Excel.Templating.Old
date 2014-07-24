<?php

namespace Excel\Templating;

class TemplatingTest extends \PHPUnit_Framework_TestCase
{
    /**
     * @test
     */
    public function load時にinitializeが実行される()
    {
        $templating = $this->getMockBuilder('\Excel\Templating\Templating')
            ->disableOriginalConstructor()
            ->setMethods(['initialize'])
            ->getMock()
        ;
        $templating
            ->expects($this->once())
            ->method('initialize')
        ;

        $templating->load(__DIR__.'/data/empty.xlsx');
    }

    /**
     * @test
     */
    public function 基本的な使い方()
    {
        $dummyTemplatePath = __DIR__.'/data/empty.xlsx';
        $dummyOutputPath = __DIR__.'/output/output.xlsx';

        $serviceFactory = $this->getMock('\Excel\Templating\ServiceFactory');
        $templating = new Templating($serviceFactory);
        $templating
          ->load($dummyTemplatePath)
          ->save($dummyOutputPath)
        ;

        $this->assertEquals(file_get_contents($dummyTemplatePath), file_get_contents($dummyOutputPath));
    }
}
