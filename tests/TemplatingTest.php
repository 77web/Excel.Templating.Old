<?php

namespace Excel\Templating;

use Excel\Templating\Service\Renderer;

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
        $dummy = [];
        $service = $this->getMockForAbstractClass('\Excel\Templating\Service\Service');

        $serviceFactory = $this->getMock('\Excel\Templating\ServiceFactory');
        $serviceFactory
            ->expects($this->once())
            ->method('create')
            ->with('test')
            ->will($this->returnValue($service))
        ;
        $service
            ->expects($this->once())
            ->method('execute')
            ->with($this->isInstanceOf('\ZipArchive'), $dummy)
        ;

        $templating = new Templating($serviceFactory);
        $templating
            ->load($dummyTemplatePath)
            ->addService('test', $dummy)
            ->save($dummyOutputPath)
        ;

        $this->assertEquals(file_get_contents($dummyTemplatePath), file_get_contents($dummyOutputPath));
    }

    /**
     * @test
     */
    public function renderショートカットによりrendererサービスを使う()
    {
        $dummyTemplatePath = __DIR__.'/data/empty.xlsx';
        $dummyOutputPath = __DIR__.'/output/output.xlsx';
        $dummyVariables = [];
        $service = $this->getMock('\Excel\Templating\Service\Renderer');

        $serviceFactory = $this->getMock('\Excel\Templating\ServiceFactory');
        $serviceFactory
            ->expects($this->once())
            ->method('create')
            ->with('renderer')
            ->will($this->returnValue($service))
        ;
        $service
            ->expects($this->once())
            ->method('execute')
            ->with($this->isInstanceOf('\ZipArchive'), $dummyVariables)
        ;

        $templating = new Templating($serviceFactory);
        $templating
          ->load($dummyTemplatePath)
          ->render($dummyVariables)
          ->save($dummyOutputPath)
        ;

        $this->assertEquals(file_get_contents($dummyTemplatePath), file_get_contents($dummyOutputPath));
    }
}
