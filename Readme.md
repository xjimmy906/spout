### Writer
```php
use iJiaXin\Writer;
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
        if($val[1] == "man"){
            $val[1] = "男";
        }elseif([$val[1] == 'woman']){
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
use iJiaXin\Reader;
use Box\Spout\Common\Type;
$reader = new Reader([
    'fileName'=>'/Users/xujw/Downloads/test.xlsx',
    'type'=>Type::XLSX,
]);
//设置读取的sheet
$reader->setReaderSheet(0);
//读取当前sheet中的数据总条数
$reader->count();
//迭代器读取每一行数据
foreach($reader->rowIterator() as $key=>$val){
    var_dump($key,$val);
}
```