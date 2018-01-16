using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Threading;

namespace ConversionTools
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public Process process = null;
        private void Form1_Load(object sender, EventArgs e)
        {
            
        }


        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog open1 = new OpenFileDialog();
            open1.Multiselect = true;//允许同时选择多个文件
            open1.Filter = "txt files(*.wav)|*.wav|All files(*.*)|*.*";
            if (this.textBox1.Text == "")
            {

                MessageBox.Show("请先选择转换后文件的存储位置");

            }
            else
            {
                if (open1.ShowDialog() == DialogResult.OK)
                {
                    int progressVarValue = 0;//进度初始值
                    for (int fi = 0; fi < open1.FileNames.Length; fi++)
                    {
                        //获得源文件的路径名称
                        string fileName = open1.FileNames[fi].ToString();//路径

                        FileInfo fi1 = new FileInfo(open1.FileNames[fi].ToString());

                        string strName = fi1.Name;

                        string targetFileName = this.textBox1.Text + "\\" + strName.Replace(".wav", ".mp3");

                        ConvertToMp3(System.Windows.Forms.Application.StartupPath, fileName, targetFileName);

                        if (fi == open1.FileNames.Length - 1)
                        {
                            progressVarValue = 100;
                            this.label3.Text = "文件已全部处理完成";
                        }
                        else
                        {

                            this.label3.Text = "正在处理文件：" + strName;
                            progressVarValue = 100 / open1.FileNames.Length * fi;
                        }

                        this.progressBar1.Value = progressVarValue;

                    }

                }

            }
            
        }


        private void button2_Click_1(object sender, EventArgs e)
        {
            try
            {

                string content = this.textBox1.Text;  //文件内容
                string path = string.Empty;  //文件路径
                FolderBrowserDialog save = new FolderBrowserDialog();

                if (save.ShowDialog() == DialogResult.OK)
                    path = save.SelectedPath;
                if (path != string.Empty)
                {
                    this.textBox1.Text = path;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }


        }

        /// <summary>
        /// 
        /// avi 转MP4  存储位置选择并回显
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {

            try
            {

                string content = this.textBox2.Text;  //文件内容
                string path = string.Empty;  //文件路径
                FolderBrowserDialog save = new FolderBrowserDialog();

                if (save.ShowDialog() == DialogResult.OK)
                    path = save.SelectedPath;
                if (path != string.Empty)
                {
                    this.textBox2.Text = path;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        /// <summary>
        /// 选择需要转为Mp4的文件，并进行转码请求
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Multiselect = true;//允许同时选择多个文件
            open.Filter = "txt files(*.avi)|*.avi|All files(*.*)|*.*";
            if (this.textBox2.Text == "")
            {

                MessageBox.Show("请先选择转换后文件的存储位置");

            }
            else
            {
                if (open.ShowDialog() == DialogResult.OK)
                {
                    int progressVarValue = 0;//进度条初始值
                    for (int fi = 0; fi < open.FileNames.Length; fi++)
                    {
                        //获得源文件的路径名称
                        string fileName = open.FileNames[fi].ToString();//路径

                        FileInfo fi1 = new FileInfo(open.FileNames[fi].ToString());

                        string strName = fi1.Name;
                        
                        string targetFileName = this.textBox2.Text + "\\" + strName.Replace(".avi", ".mp4");

                        ConvertToMp4(System.Windows.Forms.Application.StartupPath, fileName, targetFileName);

                        if (fi == open.FileNames.Length - 1)
                        {
                            progressVarValue = 100;
                            this.label8.Text = "文件已全部处理完成";
                        }
                        else
                        {
                            this.label8.Text = "正在处理文件：" + strName;
                            progressVarValue = 100 / open.FileNames.Length * fi;
                        }

                        this.progressBar2.Value = progressVarValue;

                    }

                }

            }
        }

/// <summary>
/// avi转mp4方法
/// </summary>
/// <param name="applicationPath"></param>
/// <param name="fileName"></param>
/// <param name="targetFilName"></param>
        public void ConvertToMp4(string applicationPath, string fileName, string targetFilName)
        {


            Process process = new Process();

            try
            {

               string inputFile = fileName;
                string outputFile = targetFilName;

                process.StartInfo.FileName = "C:\\Windows\\Temp\\ffmpeg.exe";  // 这里也可以指定ffmpeg的绝对路径
                process.StartInfo.Arguments = " -i " + inputFile + " -pix_fmt yuv420p -y -c:v libx264 -c:a libfdk_aac -movflags faststart " + outputFile;
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.CreateNoWindow = true;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.RedirectStandardInput = true;
                process.StartInfo.RedirectStandardError = true;
                process.ErrorDataReceived += new DataReceivedEventHandler(Output);  // 捕捉ffmpeg.exe的错误信息

                DateTime beginTime = DateTime.Now;

                process.Start();
                process.BeginErrorReadLine();   // 开始异步读取


                process.WaitForExit();    // 等待转码完成

                if (process.ExitCode == 0)
                {
                    int exitCode = process.ExitCode;
                    DateTime endTime = DateTime.Now;
                    TimeSpan t = endTime - beginTime;
                    double seconds = t.TotalSeconds;
                }
                // ffmpeg.exe 发生错误
                else
                {

                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                process.Close();
            }


           
        }
/// <summary>
/// 音频文件转mp3方法
/// </summary>
/// <param name="applicationPath"></param>
/// <param name="fileName"></param>
/// <param name="targetFilName"></param>
        public void ConvertToMp3(string applicationPath, string fileName, string targetFilName)
        {
            Process process = new Process();

            try
            {

                string inputFile = fileName;
                string outputFile = targetFilName;

                process.StartInfo.FileName = "C:\\Windows\\Temp\\ffmpeg.exe";  // 这里也可以指定ffmpeg的绝对路径
                process.StartInfo.Arguments = " -i " + inputFile + " -y  -acodec aac " + outputFile;
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.CreateNoWindow = true;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.RedirectStandardInput = true;
                process.StartInfo.RedirectStandardError = true;
                process.ErrorDataReceived += new DataReceivedEventHandler(Output);  // 捕捉ffmpeg.exe的错误信息

                DateTime beginTime = DateTime.Now;

                process.Start();
                process.BeginErrorReadLine();   // 开始异步读取


                process.WaitForExit();    // 等待转码完成

                if (process.ExitCode == 0)
                {
                    int exitCode = process.ExitCode;
                    DateTime endTime = DateTime.Now;
                    TimeSpan t = endTime - beginTime;
                    double seconds = t.TotalSeconds;
                }
                // ffmpeg.exe 发生错误
                else
                {

                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                process.Close();
            }


        }
        private static void Output(object sendProcess, DataReceivedEventArgs output)
        {
            
            if (!string.IsNullOrEmpty(output.Data))

            {
                Console.WriteLine("ffmpeg发生错误的时候才输出信息" + output.Data);
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {

        }
        /// <summary>
        /// 全部转码时的选择文件路径
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            
            try
            {

                string content = this.textBox3.Text;  //文件内容
                string path = string.Empty;  //文件路径
                FolderBrowserDialog save = new FolderBrowserDialog();

                if (save.ShowDialog() == DialogResult.OK)
                    path = save.SelectedPath;
                if (path != string.Empty)
                {
                    this.textBox3.Text = path;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        /// <summary>  
        /// 列出指定目录下及所其有子目录及子目录里更深层目录里的文件（需要递归）  
        /// </summary>  
        /// <param name="path"></param>  
        public static List<FileInfo> GetAllFiles(string path)
        {
            DirectoryInfo dir = new DirectoryInfo(path);

            //找到该目录下的文件  
            //FileInfo[] fi = dir.GetFiles();
            FileInfo[] fi = dir.GetFiles().Where((FileInfo f) => new string[] { ".wav",".wma",".rm",".midi",".ape",".flac",".mp3",".mp4",
                ".avi",".mpg",".mpeg",".avi",".rm",".rmvb",".mov",".wmv",".asf",".dat(VCD)",".asx",".wvx",".mpe",".mpa" }.Contains(f.Extension)).ToArray();
            //把FileInfo[]数组转换为List  
            List<FileInfo> list = fi.ToList<FileInfo>();
            //foreach (FileInfo f in fi)
            //{
            //   Console.WriteLine("完整路径：" + f.FullName.ToString() + " 文件名：" + f.Name);

            //}
            
            //找到该目录下的所有目录再递归 
            DirectoryInfo[] subDir = dir.GetDirectories();

            foreach (DirectoryInfo d in subDir)

            {
                list.AddRange(GetAllFiles(d.FullName));
            }
            return list;
        }
/// <summary>
/// 获得所有音频和视频的文件名
/// </summary>
/// <param name="sender"></param>
/// <param name="e"></param>
        private void button6_Click(object sender, EventArgs e)
        {


            if (this.textBox3.Text == "")
            {

                MessageBox.Show("请先选择需要转换文件的根目录");

            }
            else
            {
                List<FileInfo> list1 = GetAllFiles(this.textBox3.Text);

                int progressVarValue = 0;//进度条初始值
                for (int i = 0; i < list1.Count; i++)
                {
                    //////////////////////////////////

                    string Errorstr = "";

                    Process p = new Process();
                    try
                    {
                        p.StartInfo.FileName = "C:\\Windows\\Temp\\ffmpeg.exe";
                        p.StartInfo.Arguments = " -i " + list1[i].FullName.ToString() + " -y ";
                        p.StartInfo.UseShellExecute = false;
                        p.StartInfo.CreateNoWindow = true;
                        p.StartInfo.RedirectStandardError = true;
                        //启动进程
                        p.Start();
                        //等待进程结束
                        p.WaitForExit();
                        Errorstr = p.StandardError.ReadToEnd();
                        string ss1 = "bitrate:";
                        string ss2 = "kb/s";
                        string kbps = Errorstr.Remove(0, Errorstr.IndexOf(ss1) + ss1.Length);
                        string kbpsstr = kbps.Substring(0, kbps.IndexOf(ss2));

                        p.Close();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    finally
                    {


                    }
                    ///////////////////////////////////

                    Console.WriteLine("完整路径：" + list1[i].FullName.ToString() + " 文件名：" + list1[i].Name);
                    //如果格式是音频的，并且格式不是MP3格式的并且后缀不是。MP3的，进入转码逻辑
                    string compareStr = Errorstr.Substring(Errorstr.IndexOf("Stream") + 1);
                    if (compareStr.Contains("Audio") && !compareStr.Contains("Audio: mp3") && !list1[i].FullName.ToString().EndsWith(".mp3"))
                    {

                        allConvertToMp3(list1[i].FullName.ToString(), list1[i].FullName.ToString().Split('.')[0]);
                       File.Delete(list1[i].FullName.ToString());
                    }
                    //如果格式是视频的，并且编码格式不是h264格式的，后缀名称也不是.mp4的视频文件，进入转码格式
                    if (compareStr.Contains("Video") && !compareStr.Contains("Video: h264") && !list1[i].FullName.ToString().EndsWith(".mp4"))
                    {
                        allConvertToMp4(list1[i].FullName.ToString(), list1[i].FullName.ToString().Split('.')[0]);
                        File.Delete(list1[i].FullName.ToString());
                    }
                    
                    if (i == list1.Count - 1)
                    {
                        progressVarValue = 100;
                        this.label11.Text = "文件已全部处理完成";
                    }
                    else
                    {
                        this.label11.Text = "正在处理文件：" + list1[i].Name;
                        progressVarValue = 100 / list1.Count * i;
                    }

                    this.progressBar3.Value = progressVarValue;
                   //
                }

            }
            
        }


        /// <summary>
        /// 音频文件转mp3方法
        /// </summary>
        /// <param name="applicationPath"></param>
        /// <param name="fileName"></param>
        /// <param name="targetFilName"></param>
        public void allConvertToMp3(string fileName,string outPath)
        {
            Process process = new Process();

            try
            {

                string inputFile = fileName;
                Console.WriteLine("inputFile：" + inputFile);
                process.StartInfo.FileName = "C:\\Windows\\Temp\\ffmpeg.exe";  // 这里也可以指定ffmpeg的绝对路径
                process.StartInfo.Arguments = " -i " + inputFile + " -y -vol 200 " + outPath+ ".mp3";
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.CreateNoWindow = true;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.RedirectStandardInput = true;
                process.StartInfo.RedirectStandardError = true;
                process.ErrorDataReceived += new DataReceivedEventHandler(Output);  // 捕捉ffmpeg.exe的错误信息

                DateTime beginTime = DateTime.Now;

                process.Start();
                process.BeginErrorReadLine();   // 开始异步读取


                process.WaitForExit();    // 等待转码完成

                if (process.ExitCode == 0)
                {
                    int exitCode = process.ExitCode;
                    DateTime endTime = DateTime.Now;
                    TimeSpan t = endTime - beginTime;
                    double seconds = t.TotalSeconds;
                }
                // ffmpeg.exe 发生错误
                else
                {

                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                process.Close();
            }


        }
/// <summary>
/// 音频文件转码为编码格式为h264后缀名为.mp4的音频文件
/// </summary>
/// <param name="applicationPath"></param>
/// <param name="fileName"></param>
/// <param name="targetFilName"></param>
        public void allConvertToMp4(string fileName, string outPath)
        {


            Process process = new Process();

            

                string inputFile = fileName;

                process.StartInfo.FileName = "C:\\Windows\\Temp\\ffmpeg.exe";  // 这里也可以指定ffmpeg的绝对路径
              process.StartInfo.Arguments = " -i " + inputFile + " -pix_fmt yuv420p -y -c:v libx264 -c:a libfdk_aac -movflags faststart " + outPath + ".mp4";
            //process.StartInfo.Arguments = " -i " + inputFile + " -y -vcodec libx264 -x264opts keyint=123:min-keyint=20 -an -c:a libfdk_aac " + outPath + ".mp4";
            //process.StartInfo.Arguments = " -i " + inputFile + " -y  -vcodec copy -acodec copy -qsame " + outPath + ".mp4";
            //process.StartInfo.Arguments = " -re -i " + inputFile + " -g 52 -acodec libvo_aacenc -ab 64k -vcodec h264 -vb 448k -f mp4 -movflags frag_keyframe+empty_moov " + outPath + ".mp4";
            //-r视频帧率 
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.CreateNoWindow = true;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.RedirectStandardInput = true;
                process.StartInfo.RedirectStandardError = true;
                process.ErrorDataReceived += new DataReceivedEventHandler(Output);  // 捕捉ffmpeg.exe的错误信息
                

                process.Start();
                process.BeginErrorReadLine();   // 开始异步读取


                process.WaitForExit();    // 等待转码完成
            
                process.Close();
                process.Dispose();//释放资源

        }
        /// <summary>
        /// 截取缩略图实现
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void interceptingThumbnails(string fileName, string outPath)
        {


            Process process = new Process();



            string inputFile = fileName;

            process.StartInfo.FileName = "C:\\Windows\\Temp\\ffmpeg.exe";  // 这里也可以指定ffmpeg的绝对路径
            process.StartInfo.Arguments = " -i " + inputFile + " -ss 0.05 -y -vcodec mjpeg " + outPath + ".png";
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.RedirectStandardInput = true;
            process.StartInfo.RedirectStandardError = true;
            process.ErrorDataReceived += new DataReceivedEventHandler(Output);  // 捕捉ffmpeg.exe的错误信息


            process.Start();
            process.BeginErrorReadLine();   // 开始异步读取


            process.WaitForExit();    // 等待转码完成

            process.Close();
            process.Dispose();//释放资源


        }
        /// <summary>
        /// 截取缩略图点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button7_Click(object sender, EventArgs e)
        {

            if (progressBar3.Value == 100 || progressBar3.Value == 0)
            {
                if (this.textBox3.Text == "")
                {

                    MessageBox.Show("请先选择需要转换文件的根目录");

                }
                else
                {
                    List<FileInfo> list1 = GetAllFiles(this.textBox3.Text);

                    int progressVarValue = 0;//进度条初始值
                    for (int i = 0; i < list1.Count; i++)
                    {
                        if (list1[i].FullName.ToString().EndsWith(".mp4"))
                        {
                            interceptingThumbnails(list1[i].FullName.ToString(), list1[i].FullName.ToString().Split('.')[0]);
                        }


                        if (i == list1.Count - 1)
                        {
                            progressVarValue = 100;
                            this.label11.Text = "文件已全部处理完成";
                        }
                        else
                        {
                            this.label11.Text = "正在处理文件：" + list1[i].Name;
                            progressVarValue = 100 / list1.Count * i;
                        }

                        this.progressBar3.Value = progressVarValue;
                        //
                    }

                }
            }
            else {
                MessageBox.Show("程序正在处理其他文件，请稍后！");
            }

        }
    }
}
