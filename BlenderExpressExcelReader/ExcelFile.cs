using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
// ReSharper disable InconsistentNaming

namespace BlenderExpressExcelReader
{
    public class ExcelFile<T>
    {
        private const int COLUMN_1 = 1;
        private const int COLUMN_2 = 2;
        private const int COLUMN_3 = 3;
        private const int COLUMN_4 = 4;
        private const int COLUMN_5 = 5;
        private const string VALOR_INVALIDO = "Valor inválido na coluna {0} - linha {1}.";
        public byte[] Conteudo { get; set; }
        public List<T> Itens { get; set; }
        private readonly Func<int, int, string, DateTime, string, T> Converter;
        public ExcelFile(byte[] conteudoArquivo,Func<int,int,string,DateTime,string,T> converter)
        {
            Conteudo = conteudoArquivo;
            Itens = Extensions.GetList<T>();
            Converter = converter;
        }
        public IList<string> CarregarExcel()
        {
            var tempPath = Path.Combine(Path.GetTempPath(), "temp.xlsx");
            File.WriteAllBytes(tempPath, Conteudo);

            var erros = new List<string>();
            var newFile = new FileInfo(tempPath);
            var i = 2; //inits on seconde line if excel file has header
            using (var pck = new OfficeOpenXml.ExcelPackage(newFile))
            {
                var ws = pck.Workbook.Worksheets.First();//Could be any tab of your choice
                var eof = false;
                while (!eof)
                {

                    eof = ws.Cells[i, COLUMN_1].Value.IsNullOrEmpty() &&
                                ws.Cells[i, COLUMN_2].Value.IsNullOrEmpty() &&
                                ws.Cells[i, COLUMN_3].Value.IsNullOrEmpty() &&
                                ws.Cells[i, COLUMN_4].Value.IsNullOrEmpty() &&
                                ws.Cells[i, COLUMN_5].Value.IsNullOrEmpty();

                    if (eof) continue;

                    var rawC1 = ws.Cells[i, COLUMN_1].Value.ToString();
                    int c1;
                    if (!int.TryParse(rawC1, out c1))
                    {
                        erros.Add(string.Format(VALOR_INVALIDO, "1", i));
                        continue;
                    }

                    var rawC2 = Convert.ToString(ws.Cells[i, COLUMN_2].Value);
                    int c2;
                    if (!int.TryParse(rawC2, out c2))
                    {
                        erros.Add(string.Format(VALOR_INVALIDO, "2", i));
                        continue;
                    }

                    var rawC3 = Convert.ToString(ws.Cells[i, COLUMN_3].Value);
                    if (rawC3.IsNullOrWhiteSpace())
                    {
                        erros.Add(string.Format(VALOR_INVALIDO, "3", i));
                        continue;
                    }


                    var rawC4 = Convert.ToString(ws.Cells[i, COLUMN_4].Value);
                    DateTime c4;
                    if (!DateTime.TryParse(rawC4, out c4))
                    {
                        erros.Add(string.Format(VALOR_INVALIDO, "4", i));
                        continue;
                    }


                    var rawC5 = Convert.ToString(ws.Cells[i, COLUMN_5].Value);
                    if (rawC5.IsNullOrWhiteSpace())
                    {
                        erros.Add(string.Format(VALOR_INVALIDO, "5", i));
                        continue;
                    }

                    //TODO: acumulate the results into a List/Array
                    Itens.Add(Converter.Invoke(c1,c2,rawC3,c4,rawC5));
                    i++;
                }
            }

            return erros;
        }
    }
}
