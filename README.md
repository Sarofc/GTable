# tabtool

导表工具，excel表格导出二进制数据文件，并生成C#代码解析数据

.net 4.7.1

## TODO

- [x] 生成/解析二进制数据
- [x] 枚举Key
- [ ] 支持双端代码生成

## Excel表规则

| 标识   | 含义                                                              |
| ------ | ----------------------------------------------------------------- |
| define | "del"不导出该字段，"enum~key"键枚举。                             |
| info   | 给策划看，也会在生成代码中作为注释                                |
| type   | 字段类型。支持 byte,int,long,float,string,bytes,ints,longs,floats |
| name   | 生成代码中的字段名称。                                            |
| value  | 每个表第一个字段必须是int类型，且不能重复，作为键来使用。         |

- 支持多表/多sheet，最后生成sheet表相关的数据文件名/代码类名。但会自动过滤包含 [Ss]heet 字符串的sheet。

## 如何使用

```bat
@echo off

"../tabtool/bin/Debug/tabtool.exe" --out_client ../tabtool.test/config/ --out_cs ../tabtool.test/ --in_excel ./excel/

pause
```

--out_client 指定导出客户端导出配置文件目录<br>
--in_excel excel文件所在的目录<br>
--out_cs 导出C#代码目录，可选<br>
<!-- --in_tbs tbs文件路径（表中用到的结构体）<br> -->
<!--   --out_server 导出服务器配置文件目录<br> -->
<!--   --out_cpp 导出C++代码目录，可选<br> -->

>>☞ 参考`tool/一键导出表.bat`的用法。<!-- 暂时不提供unity示例工程。 --></br>
 ☞ Excel格式参考`tool/excel`下的表。

## Nuget依赖

- [NPOI](https://github.com/dotnetcore/NPOI)
