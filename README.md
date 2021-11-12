# UnityExcelTool

## 支持类型
`json`
## 格式
默认第一行为字段名，第二行字段类型，第三行中文注释
![](media/16366994539176.jpg)


## 字段注释注意点
如果字段不想被输出  需要在字段key前面加`#`
字典的key不能为int
![](media/16366994033353.jpg)
![](media/16366830892100.jpg)

## 字段名命名规范
* `int 12`
* `string av`
* `List|string f1,f2`
* `Dictionary|string,string  s:aa|d:bb|f:cc`
* `bool  true`

## 转换文件存放位置
* 默认Assets/Excel， 可自行修改
## 输出地方
* 支持设置CS文件及Json文件输出位置，如未设置默认Assets/Excel


## example
excel: [ExampleData22](media/ExampleData22.xlsx)
json: [ExampleData22](media/ExampleData22.json)
cs: [ExampleData22](media/ExampleData22.cs)
