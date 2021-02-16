# tabtool

导表工具，excel表格导出二进制数据文件，并生成C#代码解析数据

.net 4.7.1

## 特性

- [x] 二进制数据，生成cs文件来读取数据
- [x] 枚举Key
- [x] 支持1~4个key
- [ ] 支持双端代码生成

## Excel表头

| 标识   | 含义                                                                           |
| ------ | ------------------------------------------------------------------------------ |
| define(后续会改为client,server) | "key"支持1~4个key,类型必须为int; "del"不导出该字段; "enum~key"键枚举。         |
| info   | 字段注释                                                                       |
| type   | 字段类型。支持 byte,int,long,float,string,bytes,ints,longs,floats,map<int,int> |
| name   | 生成代码中的字段名称。                                                         |
| value  | 每个表第一个字段必须是int类型，且不能重复，作为键来使用。                      |

- 支持多表/多sheet，最后生成sheet表相关的数据文件名/代码类名。但会自动过滤包含 [Ss]heet 字符串的sheet。

## 如何使用

### 1.一键打表

--out_client 指定导出客户端导出配置文件目录<br>
--in_excel excel文件所在的目录<br>
--out_cs 导出C#代码目录，可选<br>
<!--   --out_server 导出服务器配置文件目录<br> -->
<!--   --out_cpp 导出C++代码目录，可选<br> -->

>>☞ 参考`tables/*.bat`的用法。</br>
 ☞ Excel格式参考`tables/excel`下的表。

### 2.读取数据

```csharp
    // 设置数据表路径
    TableCfg.s_TableSrc = k_ConfigPath;

    // 数据表加载委托
    TableCfg.s_BytesLoader = path =>
    {
        using (var fs = new FileStream(path, FileMode.Open))
        {
            var data = new byte[fs.Length];
            fs.Read(data, 0, data.Length);
            return data;
        }
    };

    // 加载指定数据表
    csvTest.Get().Load();
    Console.WriteLine(csvTest.Get().PrintTable());

    csvTest1.Get().Load();
    Console.WriteLine(csvTest1.Get().PrintTable());

    // 通过 csvXXX.Query(key1,key2,...) 获取行数据
    Console.WriteLine(string.Join(",", csvTest2.Query(0, 0, 0).float_arr));

    // 卸载指定数据表
    csvTest2.Get().Unload();
```

## Nuget依赖

- [NPOI](https://github.com/dotnetcore/NPOI)
