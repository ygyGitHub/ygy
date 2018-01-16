using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.IO;
using System.Threading;

namespace getFilePath
{
    class Program
    {
        static void Main(string[] args)
        {
            string exe = System.IO.Directory.GetCurrentDirectory();
            List<FileInfo> l = GetAllFiles(exe);
            StreamWriter sw = new StreamWriter(@exe + "\\pathLog.txt");
            Console.SetOut(sw);
            string path = "";
            foreach (FileInfo f in l) {

                //Console.WriteLine("完整路径：" + f.FullName.ToString() + " 文件名：" + f.Name);

                string parentPath =  f.FullName.ToString().Substring(0, f.FullName.ToString().LastIndexOf("\\"));
                if (!path.Contains(parentPath)) {

                    path += f.FullName.ToString().Substring(0, f.FullName.ToString().LastIndexOf("\\"))+ "\r\n";

                }
                

            }
            Console.WriteLine(path);
            sw.Flush();
            sw.Close();
        }

        public static List<FileInfo> GetAllFiles(string path)
        {
            DirectoryInfo dir = new DirectoryInfo(path);

            //找到该目录下的文件  
            //FileInfo[] fi = dir.GetFiles();
            FileInfo[] fi = dir.GetFiles().Where((FileInfo f) => new string[] {".mp4",
                ".avi",".mpg",".mpeg",".avi",".rm",".rmvb",".mov",".wmv",".asf",".dat(VCD)",".asx",".wvx",".mpe",".mpa" }.Contains(f.Extension)).ToArray();
            //把FileInfo[]数组转换为List  
            List<FileInfo> list = fi.ToList<FileInfo>();

            //找到该目录下的所有目录再递归 
            DirectoryInfo[] subDir = dir.GetDirectories();

            foreach (DirectoryInfo d in subDir)

            {
                list.AddRange(GetAllFiles(d.FullName));
            }
            return list;
        }



    }
}
