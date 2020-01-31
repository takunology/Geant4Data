using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace G4DataConvert
{
    public class Output : AverageCulc
    {
        private static string ResultFileName;
        public void Result()
        {
            try
            {
                ExcelOut(Program.ReadDir);
                Console.WriteLine("--------------------------------------------------------");
                Console.WriteLine("\n結果が出力されました.");
                Console.WriteLine($"保存先 : {ResultFileName}");
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
        }

        private void ExcelOut(string SavePath)
        {
            string ResultDir = $@"{SavePath}\RESULT";
            Directory.CreateDirectory(ResultDir); //最終出力ディレクトリ

            IWorkbook Book;
            ResultFileName = $@"{ResultDir}\result.xlsx";
            Book = new XSSFWorkbook(); //ブック生成
            ISheet Sheet = Book.CreateSheet("Sheet1"); //シート生成
            FileStream fileStream = new FileStream(ResultFileName, FileMode.Create, FileAccess.Write);
            WriteData(Sheet);
            Book.Write(fileStream);
        }

        private void WriteData(ISheet Sheet)
        {
            try
            {
                for (int row = 0; row < AverageValue.Count; row++)
                {
                    //拡張子無しのファイル名
                    string FileName = Regex.Replace(ExcelFileName[row], ".txt.xlsx", "");
                    ExcelWriteString(Sheet, row, 0, FileName); //ファイル名を１列目に記述
                    ExcelWriteValue(Sheet, row, 1, AverageValue[row]); //各ファイルごとの平均値を2列目に記述
                    Console.WriteLine($"{FileName} : {AverageValue[row]}");
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
        }

        //ファイル名書き込み
        private void ExcelWriteString(ISheet Sheet, int ROW, int COL, string Value)
        {
            var row = Sheet.GetRow(ROW) ?? Sheet.CreateRow(ROW);
            var cell = row.GetCell(COL) ?? row.CreateCell(COL);
            cell.SetCellValue(Value);
        }

        //平均値の書き込み
        private void ExcelWriteValue(ISheet Sheet, int ROW, int COL, double Value)
        {
            var row = Sheet.GetRow(ROW) ?? Sheet.CreateRow(ROW);
            var cell = row.GetCell(COL) ?? row.CreateCell(COL);
            cell.SetCellValue(Value);
        }
    }
}
