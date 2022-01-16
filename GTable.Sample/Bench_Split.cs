using BenchmarkDotNet.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Saro.Table.sample
{
    public class Bench_Split
    {
        [Benchmark]
        public void span_split()
        {
            var line = "1,1,1,1";

            Span<byte> bytes = stackalloc byte[32];

            var span = line.AsSpan();
            var arr = span.Split(',');
            ushort index = 0;
            foreach (var item in arr)
            {
                var chrSpan = span[item].Trim();

                if (!byte.TryParse(chrSpan, out var val))
                { }
                bytes[index++] = val;
            }

            for (int k = 0; k < index; k++)
            {
                //Console.WriteLine(bytes[k]);
            }
            bytes.Clear();
        }

        [Benchmark]
        public void string_split()
        {
            var line = "1,1,1,1";

            var arr = line.Split(',');
            for (ushort i1 = 0; i1 < arr.Length; i1++)
            {
                var res = byte.TryParse(arr[i1].AsSpan().Trim(), out byte val);
            }
        }
    }
}
