<?php

namespace Excel\Templating;

use Excel\Templating\Service\Renderer;

class TemplatingTest extends \PHPUnit_Framework_TestCase
{
    /**
     * @test
     */
    public function functional_test()
    {
        $outputPath = __DIR__.'/output/functional.xlsx';
        $services = [
            'renderer' => '\Excel\Templating\Service\Renderer',
            'sheet_remover' => '\Excel\Templating\Service\SheetRemover',
            'row_remover' => '\Excel\Templating\Service\RowRemover',
            'row_concealer' => '\Excel\Templating\Service\RowConcealer',
        ];
        $serviceFactory = new ServiceFactory($services);
        $templating = new Templating($serviceFactory);

        $templating
            ->load(__DIR__.'/data/template.xlsx')
            ->render(['%foo%' => 'bar'])
            ->removeSheet(['Sheet2'])
            ->addService('row_concealer', ['Sheet1' => [5]])
            ->addService('row_remover', ['Sheet1' => [2, 3]])
            ->save($outputPath)
        ;

        $this->assertTrue(file_exists($outputPath));
    }

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
