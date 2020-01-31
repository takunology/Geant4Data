using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace G4DataConvert
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("読み込みたいデータのあるディレクトリを入力してください.");
            //string directory = Console.ReadLine();
            string ReadDir = @"C:\Users\takun\Desktop\G4";
            Console.WriteLine("--------------------------------------------------------");

            TextToExcel TextToExcel = new TextToExcel(); //テキスト形式からエクセル形式へ変換
            string ExcelDir = $@"{ReadDir}\EXCEL";
            Directory.CreateDirectory(ExcelDir); //Excel変換作業用ディレクトリ生成
            TextToExcel.TxtToExcel(ReadDir);

            Console.ReadKey();
        }
    }
}
