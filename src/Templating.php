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
     * 今回の変更で使用するサービスの、識別名 => 実行時引数
     */
    private $services;

    public function __construct(ServiceFactory $serviceFactory)
    {
        $this->serviceFactory = $serviceFactory;
    }

    /**
     * 使用するテンプレート名を保存し、変更内容を初期化する
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
     * 状態（変更内容）を初期化する
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
     * @param array|null $arguments
     * @return Templating
     */
    public function addService($name, array $arguments = null)
    {
        $this->services[$name] = $arguments;

        return $this;
    }

    /**
     * @param string $outputPath
     */
    public function save($outputPath)
    {
        copy($this->templatePath, $outputPath);

        $output = new \ZipArchive;
        $output->open($outputPath);

        foreach ($this->services as $serviceName => $arguments) {
            $service = $this->serviceFactory->create($serviceName);
            $service->execute($output, $arguments);
        }

        $output->close();
    }
}
