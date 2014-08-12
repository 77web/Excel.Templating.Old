<?php


namespace Excel\Templating;

/**
 * Class ServiceFactory
 * create service
 *
 * @package Excel\Templating
 */
class ServiceFactory
{
    const COLUMN_CONCEALER = 'column_concealer';
    const COLUMN_REMOVER = 'column_remover';
    const RENDERER = 'renderer';
    const ROW_CONCEALER = 'row_concealer';
    const ROW_REMOVER = 'row_remover';
    const SHEET_REMOVER = 'sheet_remover';

    /**
     * @var array
     * an assoc with service name => FQCN
     */
    private static $coreServiceList;

    /**
     * @var array
     * an assoc with service name => FQCN
     */
    private $availableServices;

    /**
     * @param array $serviceList
     */
    public function __construct(array $serviceList = [])
    {
        self::initializeCoreServices();

        $this->availableServices = array_merge(self::$coreServiceList, $serviceList);
    }

    /**
     * initialize core service
     */
    private static function initializeCoreServices()
    {
        self::$coreServiceList = [
            self::COLUMN_CONCEALER => '\Excel\Templating\Service\ColumnConcealer',
            self::COLUMN_REMOVER => '\Excel\Templating\Service\ColumnRemover',
            self::RENDERER => '\Excel\Templating\Service\Renderer',
            self::ROW_CONCEALER => '\Excel\Templating\Service\RowConcealer',
            self::ROW_REMOVER => '\Excel\Templating\Service\RowRemover',
            self::SHEET_REMOVER => '\Excel\Templating\Service\SheetRemover',
        ];
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
