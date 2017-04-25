using System;
using System.Configuration;
using System.IO;
using Excel.Model;

namespace BlenderExpressExcelReader
{
    class Program
    {
        static void Main(string[] args)
        {

            var fileName = ConfigurationManager.AppSettings["fileName"];
            var isLocalFile =  ConfigurationManager.AppSettings["isLocalFile"].To<bool>();
            var directory = isLocalFile ? Environment.CurrentDirectory : ConfigurationManager.AppSettings["directory"];

            using (var fs = File.Open($"{directory}/{fileName}", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var bs = new BufferedStream(fs))
            using (var sr = new StreamReader(bs))
            {
                var xfile = new ExcelFile<Dummy>(Get(sr), GetDto);
                xfile.CarregarExcel();
                xfile.Itens.ForEach(x =>
                {
                    Console.WriteLine($"{x.C1}\t{x.C2}\t{x.C3}\t{x.C4}\t{x.C5}\t");
                });
            }
            Console.ReadKey();
        }

        static Dummy GetDto(int c1, int c2, string c3, DateTime c4, string c5)
        {
            return new Dummy()
            {
                C1 = c1,
                C2 = c2,
                C3 = c3,
                C4 = c4,
                C5 = c5
            };   
        }
        static byte[] Get(StreamReader reader)
        {
            byte[] bytes;
            using (var memstream = new MemoryStream())
            {
                reader.BaseStream.CopyTo(memstream);
                bytes = memstream.ToArray();
            }
            return bytes;
        }
    }
}
