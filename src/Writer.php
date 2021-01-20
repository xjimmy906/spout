<?php

namespace iJiaXin;


use Box\Spout\Common\Type;
use Box\Spout\Writer\Style\Color;
use Box\Spout\Writer\Style\Style;
use Box\Spout\Writer\Style\StyleBuilder;
use Box\Spout\Writer\WriterFactory;
use Box\Spout\Writer\WriterInterface;

class Writer extends BaseObject
{

    /**
     * @var WriterInterface
     */
    public $writer = null;

    /**
     * @var string 导入或导出的文件格式
     */
    public $type = Type::XLSX;

    /**
     * @var string 导入或导出的文件路径名
     */
    public $fileName = "";

    /**
     * @var bool 是否是下载
     */
    public $download = false;

    /**
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    public function init()
    {
        parent::init();
        $this->writer = WriterFactory::create($this->type);
        if($this->type === Type::CSV) {
            $this->writer->setShouldAddBOM(false);
        }
        if($this->download === true) {
            $this->writer->openToBrowser($this->fileName);
        }else{
            $this->writer->openToFile( $this->fileName );
        }
    }

    /**
     * @param array      $header
     * @param Style|null $style
     *
     * @return $this
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    public function setHeader(array $header,Style $style = null)
    {
        if($style === null) {
            $style = (new StyleBuilder())
                ->setFontBold()
                ->setFontSize(20)
                ->setFontColor(Color::BLACK)
                ->setShouldWrapText()
                ->build();
        }
        $this->addRow($header,$style);
        return $this;
    }

    /**
     * @param array         $data
     * @param Style|null    $style
     * @param callable|null $callable
     *
     * @return $this
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    public function addRow(array $data,Style $style = null,callable $callable = null)
    {
        if (is_callable($callable)) {
            $data = $callable($data,$this->writer);
            if($data === false){
                return $this;
            }
        }
        if($style){
            $this->writer->addRowWithStyle($data,$style);
        }else{
            $this->writer->addRow($data);
        }
        return $this;
    }

    /**
     * @param array         $data
     * @param Style|null    $style
     * @param callable|null $callable
     *
     * @return $this
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    public function addRows(array $data,Style $style = null,callable $callable = null)
    {
        if (is_callable($callable)) {
            $data = $callable($data,$this->writer);
            if($data === false){
                return $this;
            }
        }
        if($style){
            $this->writer->addRowsWithStyle($data,$style);
        }else{
            $this->writer->addRows($data);
        }
        return $this;
    }

    /**
     * 释放资源
     */
    public function finish(){
        if($this->download === true){
            $this->writer->close();
            die();
        }
        $this->reader->close();
    }

    /**
     * destruct
     */
    public function __destruct() {
        $this->finish();
    }
}