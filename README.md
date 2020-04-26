# tabtool

开发中...

导表工具，excel表格导出csv配置文件并生成C#代码解析配表

## TODO

- [ ] 生成/解析二进制数据
- [ ] 枚举代码
- [ ] 支持双端代码生成

## excel表规则

- 第一行filter，"client"标识该字段只导出到客户端配置文件，"server"标识只导出到服务器配置文件，"all"标识前后端都需要，"not"标识不导出该字段。
- 第二行注释，给策划看，也会在生成代码中作为注释
- 第三行name，也是生成代码中的结构体字段名称
- 第四行type，参考下面的字段类型说明。
- 每个表第一个字段必须是id字段，id必须从1开始，0是读表错误。
- id字段的filter如果标识为"client"则表示这个表只导出客户端配置文件，不导出服务器配表。反之亦然。
- 支持多表/多sheet，最后生成sheet表明相关的数据文件名/代码类名。但会自动过滤包含 [Ss]heet 字符串的sheet

 ## 字段类型

- int 整数和bool
- float 浮点数
- string 字符串
- int+ 整数迭代
- float+ 浮点数迭代
<!-- - string+ 字符串迭代 -->
<!-- - tbsIdCount 定义在meta.tbs中的结构体
- tbsIdCount+ 结构体迭代 -->
    
## 复合字段及其迭代

- 一级字段迭代：`11,22,33,44`在type中用`int+`表示。
<!-- - 二级字段迭代：`1,1;2,2`在type中用`tbsIdCount+`表示。
- 通过一个结构描述文件支持结构体，`meta.tbs`。
- 我认为表字段结构体嵌套是没有意义的，所以仅支持到二级复合字段。
- 注意excel中填写,时要设置单元格为文本模式，否则会变成数字分隔符。
- tbs文件非常简单，如下就定义一个结构体tbsIdCount:

```c
//表示id和数量
tbsIdCount {
    id int
    count int
}
``` -->

## 代码生成

- C#版本  tbs文件生成TableStruct.cs  csv文件生成TableConfig.cs
<!-- - C++版本 tbs文件tablestruct.h  csv文件生成生成一对tableconfig.h/tableconfig.cpp -->
<!-- - Go版本  TODO 暂时没用到，用到了再支持 -->

## 错误检查

类型模式不匹配的字段会在打表过程中检查出来。<br>
`目前仅采用 \t 分割csv文件，可能存在未知问题，未测试`

## 导表工具使用

>> ☞ 参考test目录中`一键导出表.bat`的用法。暂时不提供unity示例工程。

```bat
@echo off

".\tools\tabtool.exe" --out_client .\data\config\ --out_cs .\data\table_cs\ --in_excel .\excel\ --in_tbs .\excel\meta.tbs

set unity_project=..\
set project_data=%unity_project%\Assets\StreamingAssets\Config
set project_cs=%unity_project%\Assets\MGF\Common\DataTable\Generate

md %project_data%
md %project_cs%

del %project_data%\*.txt
copy .\data\config\*.txt %project_data% >nul 2>nul

del %project_cs%\*.cs
copy .\data\table_cs\*.cs %project_cs% >nul 2>nul

pause
```

--out_client 指定导出客户端导出配置文件目录<br>
--in_excel excel文件所在的目录<br>
--in_tbs tbs文件路径（表中用到的结构体）<br>
--out_cs 导出C#代码目录，可选<br>
<!--   --out_server 导出服务器配置文件目录<br> -->
<!--   --out_cpp 导出C++代码目录，可选<br> -->

## 第三方依赖

- [NPOI](https://github.com/dotnetcore/NPOI)
