using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace G4DataConvert
{
    public class TextToExcel
    {
        private Dictionary<int, string> txtFileName = new Dictionary<int, string>(); //テキストファイル名
        private Dictionary<int, string> txtFilePath = new Dictionary<int, string>(); //テキストファイル絶対パス
        public static string ExcelDir; //Excelファイルの絶対パス
        public static string[] TxtLines; //行数
        public static string[] Value; //列数

        //指定したディレクトリ内のTxt形式ファイルを確認する
        public void TxtToExcel (string TxtFilePath)
        {
            try
            {
                SearchTxtFiles(TxtFilePath);
                Console.WriteLine("EXCELファイルへ変換中... ");
                for(int i = 0; i <  txtFilePath.Count; i++)
                {
                    Console.Write("{0} / {1}", i + 1, txtFileName.Count + 1);
                    Console.SetCursorPosition(0, Console.CursorTop);
                    MakeExcelFile(i); //i番目のDictionaryを参照してエクセルデータ作成
                    //Thread.Sleep(200);
                }
                Console.WriteLine("ファイル変換が完了しました.");
                Console.WriteLine("--------------------------------------------------------");

            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
        }

        private void SearchTxtFiles(string TxtFilePath)
        {
            try
            {
                int dickey = 0; //Dictionary用キー
                DirectoryInfo directoryInfo = new DirectoryInfo(TxtFilePath);
                FileInfo[] FileCount = directoryInfo.GetFiles("*.txt", SearchOption.TopDirectoryOnly);
                foreach (FileInfo txtfile in FileCount)
                {
                    txtFileName.Add(dickey, txtfile.Name);
                    txtFilePath.Add(dickey, txtfile.FullName);
                    //ファイル存在確認
                    Console.WriteLine(txtFileName[dickey]);
                    dickey++;
                }
                Console.WriteLine($"\n{dickey + 1} 個のファイルが見つかりました.");
                Console.WriteLine("--------------------------------------------------------");
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
        }

        //Excelファイル作成
        private void MakeExcelFile(int DicCount)
        {
            try
            {
                ExcelDir = $@"{Program.ReadDir}\EXCEL";
                Directory.CreateDirectory(ExcelDir); //Excel変換作業用ディレクトリ生成

                IWorkbook Book;
                string ExcelFileName = $@"{ExcelDir}\{txtFileName[DicCount]}.xlsx";
                Book = new XSSFWorkbook(); //ブック生成
                ISheet Sheet = Book.CreateSheet("Sheet1"); //シート生成
                FileStream fileStream = new FileStream(ExcelFileName, FileMode.Create, FileAccess.Write);
                WriteData(DicCount, Sheet);
                Book.Write(fileStream);
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
        }

        //テキストファイル読み込みと書き込み
        private void WriteData(int DicCount, ISheet Sheet)
        {
            try
            {
                //ファイルの行数を取得
                TxtLines = File.ReadAllLines(txtFilePath[DicCount]);
                int lineCount = 0; //行数識別用変数
                //ヘッダー読み込み
                StreamReader streamReader = new StreamReader(txtFilePath[DicCount]);
                //１行ずつタブで区切ってから書き込んでいく
                while(streamReader.EndOfStream == false)
                {
                    string TmpChar = streamReader.ReadLine();
                    if (lineCount == 0) //先頭だけカンマ区切り
                    {
                        TmpChar = Regex.Replace(TmpChar, "[#\" ]", ""); //#を除く
                        string[] Header = TmpChar.Split(',');
                        for(int col = 0; col < Header.Length; col++)
                        {
                            ExcelWriteString(Sheet, 0, col, Header[col]);
                        }
                    }
                    else // データ部分はタブ区切り
                    {
                        //TmpChar = Regex.Replace(TmpChar, "[\n ]", ""); //#を除く
                        Value = TmpChar.Split('\t');
                        //for(int row = 0; row < TxtLines.Length; row++)
                        {
                            for (int col = 0; col < Value.Length; col++)
                            {
                                if (Regex.IsMatch(Value[col], "[0-9.]")) //最後はnullなので飛ばす
                                    ExcelWriteValue(Sheet, lineCount, col, double.Parse(Value[col]));
                            }
                        }
                    }
                    lineCount++;
                }

            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
        }

        //ヘッダー書き込み
        private void ExcelWriteString(ISheet Sheet, int ROW, int COL, string Value)
        {
            var row = Sheet.GetRow(ROW) ?? Sheet.CreateRow(ROW);
            var cell = row.GetCell(COL) ?? row.CreateCell(COL);
            cell.SetCellValue(Value);
        }

        //値の書き込み
        private void ExcelWriteValue(ISheet Sheet, int ROW, int COL, double Value)
        {
            var row = Sheet.GetRow(ROW) ?? Sheet.CreateRow(ROW);
            var cell = row.GetCell(COL) ?? row.CreateCell(COL);
            cell.SetCellValue(Value);
        }
    }
}
