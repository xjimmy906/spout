<?php
namespace iJiaXin;

use Box\Spout\Common\Type;
use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;
use Box\Spout\Reader\ReaderFactory;
use Box\Spout\Reader\ReaderInterface;
use Box\Spout\Reader\XLSX\Sheet;

class Reader extends BaseObject
{
    /**
     * @var \Box\Spout\Reader\CSV\Reader|\Box\Spout\Reader\ODS\Reader|\Box\Spout\Reader\XLSX\Reader $reader
     */
    public $reader = null;

    /**
     * @var string 导入或导出的文件格式
     */
    public $type = Type::XLSX;

    /**
     * @var string 导入或导出的文件路径名
     */
    public $fileName = "";

    /**
     * @var Sheet|\Box\Spout\Reader\CSV\Sheet|\Box\Spout\Reader\ODS\Sheet $sheet
     */
    public $sheet = null;

    /**
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     */
    public function init()
    {
        parent::init();
        $this->reader = ReaderEntityFactory::createReaderFromFile($this->fileName);
        $this->reader->setShouldFormatDates(true);
        $this->reader->open($this->fileName);
        $this->setReaderSheet(0);
    }

    /**
     * 迭代器获取一行数据
     * @return \Box\Spout\Reader\CSV\RowIterator|\Box\Spout\Reader\ODS\RowIterator|\Box\Spout\Reader\XLSX\RowIterator
     */
    public function rowIterator(){
        return $this->sheet->getRowIterator();
    }

    /**
     * 获取sheet中数据条数
     * @return int
     */
    public function count(){
        return iterator_count($this->sheet->getRowIterator());
    }


    /**
     * 返回所有sheet
     * getAllSheet
     * @return \Iterator
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     * datetime
     */
    public function getAllSheet(){
        return $this->reader->getSheetIterator();
    }

    /**
     * 获取读取的sheet
     * @param string|integer $value sheet名或sheet的索引
     * @param string $by name:通过sheet名获取  index:通过sheet的索引获取
     *
     * @return $this
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException|\Exception
     */
    public function setReaderSheet($value,$by = "index"){
        if(!in_array($by,['index','name'])){
            throw new \Exception("by参数取值限定为index、name");
        }
        /**
         * @var Sheet|\Box\Spout\Reader\CSV\Sheet|\Box\Spout\Reader\ODS\Sheet $sheet
         */
        foreach ($this->getAllSheet() as $sheet){
            $action = "get".ucfirst($by);
            if($sheet->$action() === $value){
                $this->sheet = $sheet;
                break;
            }
        }
        return $this;
    }
}