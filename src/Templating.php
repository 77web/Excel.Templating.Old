<?php

namespace Excel\Templating;

class Templating
{
    /**
     * @var ServiceFactory
     */
    private $serviceFactory;

    /**
     * @var string
     */
    private $templatePath;

    /**
     * @var array
     * an assoc with service name => arguments
     */
    private $services;

    /**
     * @param ServiceFactory $serviceFactory
     */
    public function __construct(ServiceFactory $serviceFactory)
    {
        $this->serviceFactory = $serviceFactory;
    }

    /**
     * save templatePath as property and initialize
     *
     * @param string $templatePath
     * @return Templating
     */
    public function load($templatePath)
    {
        $this->templatePath = $templatePath;
        $this->initialize();

        return $this;
    }

    /**
     * discard changes
     *
     * @return Templating
     */
    public function initialize()
    {
        $this->services = [];

        return $this;
    }

    /**
     * @param string $name
     * @param mixed $argument
     * @return Templating
     */
    public function addService($name, $argument = null)
    {
        $this->services[$name] = $argument;

        return $this;
    }

    /**
     * @param string $outputPath
     */
    public function save($outputPath)
    {
        copy($this->templatePath, $outputPath);

        $delayedServices = [];
        foreach ($this->services as $serviceName => $argument) {
            // remover must be executed after concealer execution
            if (false !== strpos($serviceName, 'remover')) {
                $delayedServices[$serviceName] = $argument;
                continue;
            }
            $this->call_service($outputPath, $serviceName, $argument);
        }

        foreach ($delayedServices as $serviceName => $argument) {
            $this->call_service($outputPath, $serviceName, $argument);
        }
    }

    /**
     * @param string $outputPath
     * @param string $serviceName
     * @param mixed $argument
     */
    private function call_service($outputPath, $serviceName, $argument)
    {
        $service = $this->serviceFactory->create($serviceName);

        $output = new \ZipArchive;
        $output->open($outputPath);

        try {
            $service->execute($output, $argument);
        } catch (\Exception $e) {
            $output->unchangeAll();
        }
        $output->close();
    }

    /**
     * a shortcut to add Renderer service
     *
     * @param array $variables
     * @return Templating
     */
    public function render(array $variables)
    {
        return $this->addService(ServiceFactory::RENDERER, $variables);
    }

    /**
     * a shortcut to add SheetRemover service
     *
     * @param array $sheetNamesToDelete
     * @return Templating
     */
    public function removeSheet(array $sheetNamesToDelete)
    {
        return $this->addService(ServiceFactory::SHEET_REMOVER, $sheetNamesToDelete);
    }
}
