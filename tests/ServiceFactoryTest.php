<?php

namespace Excel\Templating;

class ServiceFactoryTest extends \PHPUnit_Framework_TestCase
{
    public function test_create()
    {
        $serviceList = [
            'test' => '\stdClass',
        ];
        $factory = new ServiceFactory($serviceList);
        $testService = $factory->create('test');

        $this->assertInstanceOf('\stdClass', $testService);
    }

    public function test_create_core()
    {
        $factory = new ServiceFactory();
        $rendererService = $factory->create('renderer');

        $this->assertInstanceOf('\Excel\Templating\Service\Renderer', $rendererService);
    }

    public function core_service_can_be_overriden_by_constructor()
    {
        $factory = new ServiceFactory([ServiceFactory::RENDERER => '\stdClass']);
        $service = $factory->create(ServiceFactory::RENDERER);

        $this->assertInstanceOf('\stdClass', $service);
    }

    /**
     * @expectedException \LogicException
     */
    public function test_create_notfound()
    {
        $serviceList = [
          'test' => '\stdClass',
        ];
        $factory = new ServiceFactory($serviceList);
        $testService = $factory->create('invalid_service_name');
    }
}
