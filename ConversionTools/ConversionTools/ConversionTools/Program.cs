using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ConversionTools
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            releaseResources();//先将ffempeg释放到本地
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }

        /// <summary>
        /// 
        /// 将ffmpeg.exe释放到C:\Windows\Temp
        /// </summary>
        private static void releaseResources()
        {
            FileStream str = new FileStream(@"C:\Windows\Temp\ffmpeg.exe", FileMode.OpenOrCreate);
            str.Write(Resource1.ffmpeg, 0, Resource1.ffmpeg.Length);
            str.Close();

        }
    }

    }