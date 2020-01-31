using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace G4DataConvert
{
    public class TextToExcel
    {
        private Dictionary<int, string> txtFileName = new Dictionary<int, string>(); //テキストファイル名
        private Dictionary<int, string> txtFilePath = new Dictionary<int, string>(); //テキストファイル絶対パス

        //指定したディレクトリ内のTxt形式ファイルを確認する
        public void TxtToExcel (string TxtFilePath)
        {
            try
            {
                SearchTxtFiles(TxtFilePath);
                Console.WriteLine("EXCEL 形式へ変換中... ");
                for(int i = 0; i <  txtFilePath.Count; i++)
                {
                    //ConvertToExcel(i); //i番目のDictionaryを参照
                }
            }
            catch(Exception ex) { Console.WriteLine(ex.ToString()); }
        }

        public void SearchTxtFiles(string TxtFilePath)
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

        //Dictionaryの要素を選ぶ
        public void ConvertToExcel(int DicCount)
        {
            
            
        }
    }
}
