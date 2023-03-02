# excel2php

### composer安装方式
~~~
composer require ayflying/excel2php
~~~

### 使用方法
~~~
use ayflying\excel2php\Load;

//获取目录下全部表格文件
$list = Load::getPath('/extend/excel2json/excel');

//转换单文件
$list = Load::getFile('/extend/excel2json/excel/test.xlsx');
~~~

