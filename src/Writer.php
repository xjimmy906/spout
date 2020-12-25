<?php

namespace xjimmy906;


use Box\Spout\Common\Entity\Style\Color;
use Box\Spout\Common\Entity\Style\Style;
use Box\Spout\Common\Type;
use Box\Spout\Writer\Common\Creator\Style\StyleBuilder;
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;

class Writer extends BaseObject
{

    /**
     * @var \Box\Spout\Writer\CSV\Writer|\Box\Spout\Writer\CSV\Writer|\Box\Spout\Writer\XLSX\Writer
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
        $this->writer = WriterEntityFactory::createWriter($this->type);
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
     * 设置表头,包含默认样式
     * @param array             $header 数据表头
     * @param Style|null|array  $style  表头样式
     *
     * @return $this
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     * @since 2020/12/24 17:52
     * @author xJimmy906
     */
    public function setHeader(array $header, $style = null)
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
     * 一次写入一行
     * @param array         $data 数据
     * @param null          $style 数据样式
     * @param callable|null $callable 回调处理数据
     *
     * @return $this
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     * @since 2020/12/24 17:50
     * @author xJimmy906
     */
    public function addRow(array $data,$style = null,callable $callable = null)
    {
        if (is_callable($callable)) {
            $data = $callable($data,$this->writer);
            if($data === false){
                return $this;
            }
        }
        if(is_array($style)){
            $cells = [];
            foreach($data as $key=>$value){
                $cells[] = WriterEntityFactory::createCell($value,$style[$key] ?? null);
            }
            $row = WriterEntityFactory::createRow($cells);
        }else {
            $row = WriterEntityFactory::createRowFromArray($data, $style);
        }
        $this->writer->addRow($row);
        return $this;
    }

    /**
     * 一次写入多行
     * @param array             $data  数据
     * @param Style|null|array  $style 数据样式
     * @param callable|null     $callable 回调处理数据
     *
     * @return $this
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     * @since 2020/12/24 17:49
     * @author xJimmy906
     */
    public function addRows(array $data, $style = null,callable $callable = null)
    {
        if (is_callable($callable)) {
            $data = $callable($data,$this->writer);
            if($data === false){
                return $this;
            }
        }
        $rows = [];
        if(is_array($style)){
            foreach($data as $value){
                $cells = [];
                foreach($value as $k=>$v) {
                    $cells[] = WriterEntityFactory::createCell($v, $style[$k] ?? null);
                }
                $rows[] = WriterEntityFactory::createRow($cells);
            }
        }else {
            foreach($data as $val){
                $rows[] = WriterEntityFactory::createRowFromArray($val,$style);
            }
        }
        $this->writer->addRows($rows);
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
        $this->writer->close();
    }

    /**
     * destruct
     */
    public function __destruct() {
        $this->finish();
    }
}