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