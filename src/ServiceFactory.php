<?php


namespace Excel\Templating;

/**
 * Class ServiceFactory
 * サービスを作る
 *
 * @package Excel\Templating
 */
class ServiceFactory
{
    /**
     * @var array
     * 識別名 => クラス名(FQCN)
     */
    private $availableServices;

    /**
     * @param array $serviceList
     */
    public function __construct(array $serviceList = null)
    {
        $this->availableServices = $serviceList;
    }

    /**
     * @param string $name
     * @return \Excel\Templating\Service\Service
     * @throws \LogicException
     */
    public function create($name)
    {
        if (!isset($this->availableServices[$name])) {
            throw new \LogicException('Service not found.');
        }

        $className = $this->availableServices[$name];

        return new $className;
    }
} 
