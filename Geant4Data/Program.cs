using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Geant4Data
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("読み込みたいデータのあるディレクトリを入力してください.");
            string dir = null;
            try //例外処理
            {
                //テスト用
                dir = @"C:\Users\takun\Desktop\test";
                //dir = Console.ReadLine();
                Console.WriteLine("--------------------------------------------------------");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }

            ExcelRead excelRead = new ExcelRead(); //読み込み用インスタンス
            try //メイン処理
            {
                excelRead.WriteFile(dir);

            }
            catch(Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }


        /// <summary>
        /// 保持されたディレクトリの中身を漁っていく
        /// エクセル内での処理として、2行目以降のデータを読み込む
        /// 5列目の平均を算出したい。
        /// </summary>
        static void Data()
        {

        }

        /// <summary>
        /// エクセルで処理されたファイルをテキストファイルに吐出す
        /// 区切りはタブで
        /// </summary>
        static void ExcelToTxt()
        {

        }
    }
}