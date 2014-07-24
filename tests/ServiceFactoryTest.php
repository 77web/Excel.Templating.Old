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
