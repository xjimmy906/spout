### Writer
基于[box/spout ](https://opensource.box.com/spout/) 的一个快速读写excel文件(csx,xlsx and ods)的php库

[https://opensource.box.com/spout/](https://opensource.box.com/spout/)

```php
use xjimmy906\Writer;
use Box\Spout\Common\Type;
$writer = new Writer([
    'fileName'=>'/Users/xujw/Downloads/test.xlsx',
    'type'=>Type::XLSX,
    'download'=>false,//true下载
]);
//设置head头，第二个参数可设置样式
$writer->setHeader(['姓名',"性别"]);
//添加一行数据，第二个参数可设置样式，第三个参数可配置一个回调函数
$writer->addRow(['jack','man']);
//添加多行数据，第二个参数可设置样式，第三个参数可配置一个回调函数
$writer->addRows([['lilei','man'],['hanmeimei','woman']]);
//使用第三参数回调处理数据
$writer->addRows([['lilei','man'],['hanmeimei','woman']],null,function($data,$w){
    foreach($data as &$val){
        if($val[1] === "man"){
            $val[1] = "男";
        }elseif([$val[1] === 'woman']){
            $val[1] = "女";
        }
    }
    unset($val);
    return $data;
});
//释放资源
$writer->finish();
```

### Reader
```php
use xjimmy906\Reader;
use Box\Spout\Common\Type;
$reader = new Reader([
    'fileName'=>'/Users/xujw/Downloads/test.xlsx',
]);
//设置读取的sheet
$reader->setReaderSheet(0);
//读取当前sheet中的数据总条数
$reader->count();
//迭代器读取每一行数据
foreach($reader->rowIterator() as $row){
    //box/spout 3.0.0+ 版本以上返回Row对象，而不是包含行值的数组(3.0.0版本以前返回的为数组)  
  
    $rowAsArray = $row->toArray();  //转化为数组,兼容2.0写法
    // OR
    $cellsArray = $row->getCells(); //这可以用来访问单元格的详细信息
}
```