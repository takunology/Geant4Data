using System;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Threading;
using NPOI.XSSF.UserModel;

namespace Geant4Data
{
    public class ExcelRead
    {
        static int SHEET = 1; //シートの枚数
        static int ROW = 0; //行数 可変長
        static int COL = 0; //列数 可変長
        public double[,,] Value = new double[SHEET, ROW, COL];

        /// <summary>
        /// 入力テキストファイルをエクセル形式に直す
        /// 入力ファイルは複数
        /// 同じディレクトリ内のテキストファイルを漁っていく
        /// 直したファイルはディレクトリを新規作成後、そこに保持
        /// </summary>
        
        public void SearchTextFile(string dir)
        {
            //そのディレクトリ内のテキストファイル
            int i = 0; //Dictionary キーの値
            DirectoryInfo directoryInfo = new DirectoryInfo(dir);
            FileInfo[] Count = directoryInfo.GetFiles("*.txt", SearchOption.TopDirectoryOnly);
            Dictionary<int, string> FileName = new Dictionary<int, string>();
            foreach(FileInfo f in Count)
            {
                FileName.Add(i,f.Name);
                i++;
            }

            foreach (string Value in FileName.Values)
            {
                Console.WriteLine(Value); //存在するテキストファイルの確認
            }

            Console.WriteLine($"\n{i + 1} 個のファイルが見つかりました.");
            Console.WriteLine("--------------------------------------------------------");
            Console.WriteLine("EXCEL 形式へ変換中... ");

            int counter = 0; //ファイルカウント用
            foreach(string Value in FileName.Values)
            {
                TextToExcel(Value); //Excel 変換
                Console.Write("{0} / {1}", counter + 1, FileName.Count + 1);
                Console.SetCursorPosition(0, Console.CursorTop);
                Thread.Sleep(200);
                counter++;
            }

            Console.WriteLine("変換が完了しました.");
            Console.ReadKey();
        }

        public void TextToExcel(string TextFileName)
        {
            try
            {
                //ディレクトリ作成
                string CreatePath = $@"C:\Users\takun\Desktop\test\EXCEL";
                Directory.CreateDirectory(CreatePath);

                IWorkbook book;
                string FileName = $@"C:\Users\takun\Desktop\test\EXCEL\{TextFileName}.xlsx";
                book = new XSSFWorkbook(); //ブック生成
                book.CreateSheet("Sheet1"); //シート生成
                FileStream fileStream = new FileStream(FileName, FileMode.Create, FileAccess.Write);
                book.Write(fileStream); //書き込み
            }
            catch(Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }

        public void ExcelOpen(string path)
        {
            var Book = WorkbookFactory.Create(path);
            for(int i = 0; i < SHEET; i++)
            {
                var Sheet = Book.GetSheetAt(i);
                for(int j = 0; j < ROW; j++)
                {
                    for(int k = 0; k < COL; k++)
                    {
                       // Value[i, j, k] = GetValue(Sheet, j, k);
                    }
                }
            }
        }

        public double GetValue(ISheet Sheet, int ROW, int COL)
        {
            double value = 0;

            return value;
        }

        public void Result()
        {

        }
    }
}
