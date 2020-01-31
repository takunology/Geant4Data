using System;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Threading;
using NPOI.XSSF.UserModel;
using System.Text.RegularExpressions;

namespace Geant4Data
{
    public class ExcelRead
    {
        static int SHEET = 1; //シートの枚数 固定
        static int ROW = 0; //行数 可変長
        static int COL = 0; //列数 可変長
        public double[,,] Value = new double[SHEET, ROW, COL];

        /// <summary>
        /// 入力テキストファイルをエクセル形式に直す
        /// 入力ファイルは複数
        /// 同じディレクトリ内のテキストファイルを漁っていく
        /// 直したファイルはディレクトリを新規作成後、そこに保持
        /// </summary>
        Dictionary<int, string> FullPath = new Dictionary<int, string>();

        public void WriteFile(string dir)
        {
            //そのディレクトリ内のテキストファイル
            int i = 0; //Dictionary キーの値
            DirectoryInfo directoryInfo = new DirectoryInfo(dir);
            FileInfo[] Count = directoryInfo.GetFiles("*.txt", SearchOption.TopDirectoryOnly);
            Dictionary<int, string> FileName = new Dictionary<int, string>();
            foreach(FileInfo f in Count)
            {
                FullPath.Add(i, f.FullName);
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
                TextToExcel(Value, counter); //Excel 変換
                Console.Write("{0} / {1}", counter + 1, FileName.Count + 1);
                Console.SetCursorPosition(0, Console.CursorTop);
                Thread.Sleep(200);
                counter++;
            }
            Console.ReadKey();

            Console.WriteLine("変換が完了しました.");
            
        }

        //引数はファイル名と絶対パス
        public void TextToExcel(string TextFileName, int counter)
        {
            try
            {
                //ディレクトリ作成
                string CreatePath = $@"C:\Users\takun\Desktop\test\EXCEL";
                Directory.CreateDirectory(CreatePath);

                //EXCEL ブック作成
                IWorkbook book;
                string FileName = $@"C:\Users\takun\Desktop\test\EXCEL\{TextFileName}.xlsx";
                book = new XSSFWorkbook(); //ブック生成
                ISheet Sheet = book.CreateSheet("Sheet1"); //シート生成
                FileStream fileStream = new FileStream(FileName, FileMode.Create, FileAccess.Write);

                // テキスト内容書き込み(先頭のみ)
                StreamReader streamReader = new StreamReader(FullPath[counter]);
                string TmpChar = streamReader.ReadLine();
                TmpChar = Regex.Replace(TmpChar, "[#\"]", ""); //#を除く
                string[] Header = TmpChar.Split(','); //カンマごとに区切る

                //ヘッダーの書き込み
                for (int i = 0; i < Header.Length; i++)
                {
                    ExcelWriteString(Sheet, 0, i, Header[i]);
                }

                //データの書き込み
                TmpChar = streamReader.ReadToEnd();
                TmpChar = Regex.Replace(TmpChar, "[#\"a-zA-Z]", ""); //文字を除く
                 //タブと改行で分ける
                string[] coldata = TmpChar.Split('\t'); 

                int x = 0;
                /*foreach(string s in coldata)
                {
                    Console.WriteLine($"{x++} の要素 {s}");
                    //Console.WriteLine(coldata.Length.ToString());
                    Thread.Sleep(200);
                }*/
                

                //テキストファイルの行数を取得
                string[] lines = File.ReadAllLines(FullPath[counter]);
                for (int i = 0; i < lines.Length; i++)
                {
                    for(int j = 0; j < 7; j++)
                    {
                        
                        //Console.WriteLine(coldata[x]);
                        ExcelWriteString(Sheet, i, j, coldata[x]);
                        x++;
                        if (x % 7 == 0) { break; }
                    }

                }




                book.Write(fileStream); //書き込み
            }
            catch(Exception e)
            {
                Console.WriteLine(e.ToString());
            }


        }



        //文字列（ヘッダー用）エクセル書き込み
        public void ExcelWriteString(ISheet Sheet, int ROW, int COL, string Value)
        {
            var row = Sheet.GetRow(ROW) ?? Sheet.CreateRow(ROW);
            var cell = row.GetCell(COL) ?? row.CreateCell(COL);
            cell.SetCellValue(Value);
        }

        //値エクセル書き込み
        public void ExcelWriteValue(ISheet Sheet, int ROW, int COL, double Value)
        {
            var row = Sheet.GetRow(ROW) ?? Sheet.CreateRow(ROW);
            var cell = row.GetCell(COL) ?? row.CreateCell(COL);
            cell.SetCellValue(Value);
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
