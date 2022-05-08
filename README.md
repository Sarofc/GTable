# GTable

导表工具，excel表格导出二进制数据文件，并生成C#代码读取数据

.net 6

## 特性

- [x] 二进制数据
- [x] 代码生成
- [x] 枚举Key
- [x] 支持1~4个key
- [ ] 客户端服务器区分
- [ ] 热更新支持（考虑使用huatuo）

## Excel表头

| 标识                            | 含义                                                                           |
| ------------------------------- | ------------------------------------------------------------------------------ |
| define(后续会改为client,server) | "key"支持1~4个key,类型必须为int; "del"不导出该字段; "enum~key"键枚举。         |
| info                            | 字段注释                                                                       |
| type                            | 字段类型。支持 byte,int,long,float,string,bytes,ints,longs,floats,map<int,int> |
| name                            | 生成代码中的字段名称。                                                         |
| value                           | 每个表第一个字段必须是int类型，且不能重复，作为键来使用。                      |

- 支持多表/多sheet，最后生成sheet表相关的数据文件名/代码类名。但会自动过滤包含 [Ss]heet 字符串的sheet。

## 性能

>> 数据表: 250个excel，每个excel，130kb，有两个sheet，每个sheet 1000行，10几列
>> 结果: 导出数据+代码，花费7000ms

## 如何使用

### 1.一键打表

--out_client 指定导出客户端导出配置文件目录<br>
--in_excel excel文件所在的目录<br>
--out_cs 导出C#代码目录，可选<br>
<!--   --out_server 导出服务器配置文件目录<br> -->
<!--   --out_cpp 导出C++代码目录，可选<br> -->

>>☞ 参考`tables/*.bat`的用法。</br>
 ☞ Excel格式参考`tables/excel`下的表。

 1. build~.bat 编译，需设置msbuild环境变量
 2. copy_for_unity~.bat 将tables目录拷贝到指定目录，需设置`UNITY_PROJECT_PATH`
 3. export_all_for_unity.bat 生成表格数据和解析代码
 4. export_data_for_unity.bat 仅生成表格数据，适用于表结构未改变的情况
 5. export~.bat 用于非unity项目，功能同export_all_for_unity.bat

### 2.读取数据

```csharp
using Saro.Table;
using System;
using System.IO;

const string k_ConfigPath = @"..\..\..\generate\data\";

// setup load handler
TableLoader.s_BytesLoader = name =>
{
    var path = k_ConfigPath + name;
    using (var fs = new FileStream(path, FileMode.Open))
    {
        var data = new byte[fs.Length];
        fs.Read(data, 0, data.Length);
        return data;
    }
};

TableLoader.s_BytesLoaderAsync = async name =>
{
    var path = k_ConfigPath + name;
    using (var fs = new FileStream(path, FileMode.Open))
    {
        var data = new byte[fs.Length];
        var buffer = new Memory<byte>(data);
        await fs.ReadAsync(buffer);
        return data;
    }
};

// load sync
{
    Console.WriteLine("load sync");
    var result = csvTest1.Get().Load();
    Console.WriteLine(csvTest1.Get().PrintTable());
    Console.WriteLine(string.Join(",", csvTest1.Query(0, 0, 0).float_arr));
    Console.WriteLine(string.Join(",", csvTest1.Query(0, 0, 0).map_int_int));
    csvTest1.Get().Unload();
}

// load async
{
    Console.WriteLine();
    Console.WriteLine("load async");
    var result = await csvTest0.Get().LoadAsync();
    Console.WriteLine(string.Join(",", csvTest0.Query(0, 0, 0).float_arr));
    Console.WriteLine(string.Join(",", csvTest0.Query(0, 0, 0).map_int_int));
    csvTest0.Get().Unload();
}

Console.ReadKey();
```

## Nuget依赖

- ExcelDataReader
