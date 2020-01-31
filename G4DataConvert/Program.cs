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
        public static string ReadDir;// = @"C:\Users\takun\Desktop\G4"; //テスト用

        static void Main(string[] args)
        {
            Console.WriteLine("読み込みたいデータのあるディレクトリを入力してください.");
            ReadDir = Console.ReadLine();         
            Console.WriteLine("--------------------------------------------------------");

            TextToExcel TextToExcel = new TextToExcel(); //テキスト形式からエクセル形式へ変換
            TextToExcel.TxtToExcel(ReadDir);

            AverageCulc averageCulc = new AverageCulc();
            averageCulc.GetAverage();

            Output output = new Output();
            output.Result();

            Console.WriteLine("\n\n --- 終了するにはキーを押してください. ---");
            Console.ReadKey();
        }
    }
}
