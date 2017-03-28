using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OCRTrst
{
    class Program
    {
        static void Main(string[] args)
        {
            for (int i = 0; i < 10; i++)
            {
                //ThreadPool.QueueUserWorkItem(new WaitCallback(ocrMuliThreTest));
                ocrMuliThreTest(new object());
            }
            Console.WriteLine("程序运行完毕");
            while (true)
            {
                Thread.Sleep(2 * 1000);
                Console.WriteLine("------------------------------------刷新缓存-----------------------------------");
            }
        }
        public static void ocrMuliThreTest(Object obj)
        {
            string path = @"C:\Users\zhensheng\OneDrive\MyDirection\GitHub\Extraction.OCR\Lib\ocrTest.png";
            string result = Extraction.OCR.ExtractionOCR.Instance.Ocr_2010(path);
            Console.WriteLine(DateTime.Now + result);
            if (string.IsNullOrEmpty(result)) Console.WriteLine("解析失败\n\n");
            else Console.WriteLine("解析成功" + "\n" + result + "\n\n");

        }
    }
}
