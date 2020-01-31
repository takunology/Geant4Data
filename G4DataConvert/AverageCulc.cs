using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace G4DataConvert
{
    public class AverageCulc : TextToExcel
    {
        public static Dictionary<int, string> ExcelFileName = new Dictionary<int, string>(); //Excelファイル名
        private Dictionary<int, string> ExcelFilePath = new Dictionary<int, string>(); //Excelファイル絶対パス
        public static Dictionary<int, double> AverageValue = new Dictionary<int, double>(); //平均値

        public void GetAverage()
        {
            SearchExcelFile(ExcelDir);
            Console.WriteLine("平均値の計算中...\n");
            for (int i = 0; i < ExcelFilePath.Count; i++)
            {
                Console.Write("{0} / {1}", i + 1, ExcelFileName.Count + 1);
                Console.SetCursorPosition(0, Console.CursorTop);
                //平均値計算
                Culc(i);
            }
            Console.WriteLine("\n計算が終了しました. 結果を出力します.");
            Console.WriteLine("--------------------------------------------------------");
        }

        private void SearchExcelFile(string ExcelPath)
        {
            try
            {
                int dickey = 0; //Dictionary用キー
                DirectoryInfo directoryInfo = new DirectoryInfo(ExcelPath);
                FileInfo[] FileCount = directoryInfo.GetFiles("*.xlsx", SearchOption.TopDirectoryOnly);
                Console.WriteLine("変換されたExcelファイルを確認しています...\n");
                foreach (FileInfo txtfile in FileCount)
                {
                    ExcelFileName.Add(dickey, txtfile.Name);
                    ExcelFilePath.Add(dickey, txtfile.FullName);
                    //ファイル存在確認
                    Console.WriteLine(ExcelFileName[dickey]);
                    dickey++;
                }
                Console.WriteLine($"\n{dickey + 1} 個のファイルが見つかりました.");
                Console.WriteLine("--------------------------------------------------------");
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
        }

        private void Culc(int dicCount)
        {
            var Book = WorkbookFactory.Create(ExcelFilePath[dicCount]); //参照先
            var Sheet = Book.GetSheetAt(0);
            double ave = 0;
            for (int row = 1; row < TxtLines.Length; row++) //ヘッダーを無視する
            {
                for (int col = 0; col < Value.Length; col++)
                {
                    if(col == 4) //5列目
                    {
                        ave += GetValue(Sheet, row, col);
                    }
                }
            }
            AverageValue[dicCount] = ave / TxtLines.Length; //平均値の保持
            Console.WriteLine($"{ExcelFileName[dicCount]} : {AverageValue[dicCount]}");
        }

        private double GetValue(ISheet Sheet, int Row, int Column)
        {
            var row = Sheet.GetRow(Row) ?? Sheet.CreateRow(Row); //例外対策(なければ空のシートを追加)
            var cell = row.GetCell(Column) ?? row.CreateCell(Column);
            double value = cell.NumericCellValue;
            return value;
        }
    }
}
