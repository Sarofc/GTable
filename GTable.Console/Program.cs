using System.Text;

namespace Saro.GTable
{
    class Program
    {
        static async Task Main(string[] args)
        {
#if NETCOREAPP1_0_OR_GREATER
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
#endif

            await TableExporter.ExportAsync(args);
        }
    }
}
