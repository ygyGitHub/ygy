using System;
using System.Collections.Generic;
using System.Linq;
using System.Text; 
using System.Threading.Tasks;
using Excel =  Microsoft.Office.Interop.Excel;
using System.IO;

namespace ConsoleApplicationExcel
{
    class Program
    {
        
        static void Main(string[] args)
        {
            string exe = System.IO.Directory.GetCurrentDirectory();
            //StreamWriter sw = new StreamWriter(@path);
            //Console.SetOut(sw);

            
            //生成对比数据源
            dataProcessingOfBaseLibrary();

            Console.WriteLine(exe);

            GetAllFiles(exe);

            if (File.Exists("d:\\辅助检查库.xlsx"))
                File.Delete("d:\\辅助检查库.xlsx");
            if (File.Exists("d:\\体格检查专项视频.xlsx"))
                File.Delete("d:\\体格检查专项视频.xlsx");
            //删除对比数据源
            
            //string time = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
            
            Console.WriteLine("执行完毕！！！按回车键退出");
            
            Console.ReadLine();
        }
        /// <summary>
        /// 获得专项体格检查视频名称地址数据
        /// </summary>
        public static void getSpecialPhysicalExaminationVideoAddress(Excel.Worksheet sheet1, int rowCount, string fileName) {


            Console.WriteLine("正在对文件名字为："+ fileName + "进行体格检查视频地址赋值操作");
            Excel.Application exceltg = new Excel.Application();//引用Excel对象

            Excel.Workbook excelBooktg = exceltg.Workbooks.Open(@"d:\\体格检查专项视频.xlsx", 0, false);//打开一个工作簿

            Excel.Worksheet sheettg;//打开工作表

            exceltg.DisplayAlerts = false;

            int sheetCounttg = excelBooktg.Sheets.Count;//一共多少个sheet；



            for (int a =1; a<= sheetCounttg; a++){



                sheettg = (Excel.Worksheet)excelBooktg.Sheets[a];




                if ("专项训练".Equals(sheettg.Name)) {

                    int rowCount_tg = sheettg.UsedRange.Rows.Count;//对比文件一共几行

                    Excel.Range rng1_tg = sheettg.Cells.get_Range("A2", "A" + rowCount_tg);
                    Excel.Range rng2_tg = sheettg.Cells.get_Range("B2", "B" + rowCount_tg);
                    Excel.Range rng3_tg = sheettg.Cells.get_Range("C2", "C" + rowCount_tg);
                    object[,] arryItem1_tg = (object[,])rng1_tg.Value2;   //get range's value
                    object[,] arryItem2_tg = (object[,])rng2_tg.Value2;
                    object[,] arryItem3_tg = (object[,])rng3_tg.Value2;
                    
                    int filePathColumnNumber = GetColumnNumbers(sheet1, "文件路径");//目标文件4
                    
                    int filePathColumnNumber_tg = GetColumnNumbers(sheettg, "文件路径");//对比文件4

                    Excel.Range rng1 = sheet1.Cells.get_Range("A2", "A" + rowCount);
                    Excel.Range rng2 = sheet1.Cells.get_Range("B2", "B" + rowCount);
                    Excel.Range rng3 = sheet1.Cells.get_Range("C2", "C" + rowCount);
                    object[,] arryItem1 = (object[,])rng1.Value2;   //get range's value
                    object[,] arryItem2 = (object[,])rng2.Value2;
                    object[,] arryItem3 = (object[,])rng3.Value2;
                    //从第二行开始，第一行是表头，上边已经取过表头所在的列号，不需要再循环中去取了
                    for (int i=2; i<= rowCount;i++) {//遍历目标文件，取出来第一列第二列第三列的值，与对比文件每一行进行比对，如果有，将对比文件的文件路径中的内容赋值给目标文件
                        if (arryItem1[i - 1, 1] == null)
                            arryItem1[i - 1, 1] = "";
                        if (arryItem2[i - 1, 1] == null)
                            arryItem2[i - 1, 1] = "";
                        if (arryItem3[i - 1, 1] == null)
                            arryItem3[i - 1, 1] = "";

                        for (int j = 2; j <= rowCount_tg; j++)
                            {
                            if (arryItem1_tg[j - 1, 1] == null)
                                arryItem1_tg[j - 1, 1] = "";
                            if (arryItem2_tg[j - 1, 1] == null)
                                arryItem2_tg[j - 1, 1] = "";
                            if (arryItem3_tg[j - 1, 1] == null)
                                arryItem3_tg[j - 1, 1] = "";
                            if (arryItem1_tg[j - 1, 1] == "" && arryItem2_tg[j - 1, 1] == "" && arryItem2_tg[j - 1, 1] == "")
                                break;
                                //如果123列与对比文件123列内容相等，进行赋值操作
                                if (arryItem1[i - 1,1].Equals(arryItem1_tg[j - 1, 1]) && arryItem2[i - 1,1].Equals(arryItem2_tg[j - 1, 1]) && arryItem3[i - 1,1].Equals(arryItem3_tg[j - 1, 1]))
                                {
                                    (sheet1.Cells[i, filePathColumnNumber]).value = (sheettg.Cells[j, filePathColumnNumber_tg]).Text.ToString().Trim();
                                (sheet1.Cells[i, filePathColumnNumber]).Font.ColorIndex = 3;
                                }

                            }
                            
                        //}
    

                    }

                }

            }
            // excel.Visible = true;
            excelBooktg.Save();

            excelBooktg.Close(false);

            excelBooktg = null;

            //退出Excel程序 
            exceltg.Quit();
            exceltg = null;

            // 10.调用GC的垃圾收集方法  
            GC.Collect();

            GC.WaitForPendingFinalizers();
            Console.WriteLine("工作表名为：" + fileName + "进行体格检查视频地址赋值操作已经完成");

        } 




        /// <summary>
        /// 预先生成两个基础数据文档
        /// </summary>
        public static void dataProcessingOfBaseLibrary()
        {
            if (File.Exists("d:\\辅助检查库.xlsx"))
                File.Delete("d:\\辅助检查库.xlsx");
            if (File.Exists("d:\\体格检查专项视频.xlsx"))
                File.Delete("d:\\体格检查专项视频.xlsx");
            byte[] Savetg = global::ConsoleApplicationExcel.Resource1.tg;
            FileStream fsObjtg = new FileStream("d:\\体格检查专项视频.xlsx", System.IO.FileMode.CreateNew);
            fsObjtg.Write(Savetg, 0, Savetg.Length);
            fsObjtg.Close();
            byte[] Savefz = global::ConsoleApplicationExcel.Resource1.fz;
            FileStream fsObjfz = new FileStream("d:\\辅助检查库.xlsx", System.IO.FileMode.CreateNew);
            fsObjfz.Write(Savefz, 0, Savefz.Length);
            fsObjfz.Close();
           //删除文件的方法 File.Delete("d:\\辅助检查库.xlsx");
        }

        /// <summary>
        /// 生命体征名称为空，复制便签的数据  A列是生命体征的，把I列的值复制给C列
        /// </summary>
        public static void copyOfTheNote(Excel.Worksheet sheet1, int rowCount, string fileName)
        {

            Console.WriteLine("正在为工作表名为：" + fileName + "体征名称赋值操作");

            string checkPoint = "检查部位分类";

            string[] inspectionArea = new string[] { "检查区域", "专项检查区域" };

            string[] node = new string[] { "思维便签填充项目", "专项便签填充项目" };

            //检查部位分类所在列
            int checkPointColumnNumber = GetColumnNumbers(sheet1, checkPoint);
            for (int i = 1; i <= rowCount; i++)
            {

                string checkPointClassification = (sheet1.Cells[i, checkPointColumnNumber]).Text.ToString().Trim();
                //如果大分类是生命体征的则继续逻辑，查看检查区域是否为空，如果为空，则将便签中的内容复制给检查区域
                if ("生命体征".Equals(checkPointClassification) && !"".Equals(checkPointClassification))
                {
                    int inspectionAreaColumnNumber = 0;

                    for (int j = 0; j < inspectionArea.Length; j++)
                    {

                        //获得检查区域所在列号
                        inspectionAreaColumnNumber = GetColumnNumbers(sheet1, (string)inspectionArea.GetValue(j));


                        if (inspectionAreaColumnNumber != 0)
                        {
                            break;
                        }

                    }
                    if (inspectionAreaColumnNumber != 0)
                    {
                        string inspectionAreaContent = (sheet1.Cells[i, inspectionAreaColumnNumber]).Text.ToString().Trim();
                        //如果检查项目为空
                        if (inspectionAreaContent == "" || inspectionAreaContent == null)
                        {
                            int nodeColumnNumber = 0;

                            for (int k = 0; k < node.Length; k++)
                            {
                                nodeColumnNumber = GetColumnNumbers(sheet1, (string)node.GetValue(k));
                                if (nodeColumnNumber != 0)
                                {
                                    break;
                                }
                            }
                            if (nodeColumnNumber != 0)
                            {
                                //将便签的值复制给检查区域
                                sheet1.Cells[i, inspectionAreaColumnNumber].value = (sheet1.Cells[i, nodeColumnNumber]).Text.ToString().Trim();
                                (sheet1.Cells[i, inspectionAreaColumnNumber]).Font.ColorIndex = 3;
                            }
                            
                        }
                        
                    }

                }

            }


            Console.WriteLine("工作表名为：" + fileName + "体征名称赋值操作已完成");


        }
        /// <summary>
        /// 
        /// 拆分血压
        /// </summary>
        public static void separationOfBloodPressure(Excel.Application excel, Excel.Worksheet sheet1, int rowCount, string fileName)
        {

            Console.WriteLine("正在对工作表名为：" + fileName + "的专项体格检查血压分成2列的操作");

            string sheetName = sheet1.Name;

            if ("专项训练".Equals(sheetName))
            {

                object misValue = Type.Missing;

                int columnNumber = GetColumnNumbers(sheet1, "项目内容1");//项目内容1 的列号

                if(!string.IsNullOrWhiteSpace((sheet1.Cells[1, columnNumber+1]).Text.ToString())) {
                    //Excel.Range range = (Excel.Range)sheet1.Columns[columnNumber + 1, misValue];

                    //range.Insert(Excel.XlDirection.xlToLeft);

                    for (int i = 1; i <= rowCount; i++)

                    {

                        string content = (sheet1.Cells[i, columnNumber]).Text.ToString().Trim();

                        if (content.Contains("mmHg"))
                        {

                            string[] contents = content.Split('/');
                            if (contents.Length>1)
                            {
                                (sheet1.Cells[i, columnNumber]).value = contents.GetValue(0) + "mmHg";

                                (sheet1.Cells[i, columnNumber]).Font.ColorIndex = 3;
                                (sheet1.Cells[i, columnNumber + 1]).value = contents.GetValue(1);
                                (sheet1.Cells[i, columnNumber + 1]).Font.ColorIndex = 3;

                            }
                           
                        }

                    }

                }



            }

            Console.WriteLine("工作表名为：" + fileName + "的专项体格检查血压分成2列的操作已经完成");

        }
        /// <summary>
        /// 将体格检查文件中的文件路径内容的avi全部替换为mp4
        /// 将脊柱、四肢第三列为空的赋值为四肢
        /// </summary>
        public static void replaceAviWithMp4(Excel.Worksheet sheet1, int rowCount, string fileName)
        {
            Console.WriteLine("正在对工作表名为：" + fileName + "的文件路径中avi替换为mp4操作");
            string colName = "文件路径";
            Console.WriteLine(sheet1.Name);
            //获得"文件路径"的列号
            int columnNumber = GetColumnNumbers(sheet1, colName);

            int checkPointColumnNumber = GetColumnNumbers(sheet1, "检查部位分类");//检查部位列号

            int checkQuYuColumnNumber = GetColumnNumbers(sheet1, "检查区域");//检查区域

            for (int i = 1; i <= rowCount; i++)
            {

                string checkPointValue = (sheet1.Cells[i, checkPointColumnNumber]).Text.ToString().Trim();
                if(checkQuYuColumnNumber>0)
                {
                    string checkQuYuValue = (sheet1.Cells[i, checkQuYuColumnNumber]).Text.ToString().Trim();
                    if ("脊柱、四肢".Equals(checkPointValue) && string.IsNullOrWhiteSpace(checkQuYuValue))
                    {
                        //将脊柱、四肢第三列为空的赋值为四肢
                        (sheet1.Cells[i, checkQuYuColumnNumber]).value = checkPointValue.Split('、').GetValue(1);
                        (sheet1.Cells[i, checkQuYuColumnNumber]).Font.ColorIndex = 3;
                    }
                }
                
                string whetherNeedReplace = (sheet1.Cells[i, columnNumber]).Text.ToString().Trim();

                if (whetherNeedReplace.Contains("avi") && whetherNeedReplace != "")
                {

                    (sheet1.Cells[i, columnNumber]).value = whetherNeedReplace.Replace(".avi", ".mp4");
                    (sheet1.Cells[i, columnNumber]).Font.ColorIndex = 3;
                }
                if (whetherNeedReplace.Contains(".wav") && whetherNeedReplace != "") {
                    (sheet1.Cells[i, columnNumber]).value = whetherNeedReplace.Replace(".wav", ".mp3");
                    (sheet1.Cells[i, columnNumber]).Font.ColorIndex = 3;

                }
            }

            Console.WriteLine("工作表名为：" + fileName + "的文件路径中avi替换为mp4操作已经完成");

        }


        /// <summary>
        /// 删除不需要的列
        /// </summary>
        public static void openExcelAndDeleteColumns(string path, string fileName)
        {
            
            Excel.Application excel = new Excel.Application();//引用Excel对象

            Excel.Workbook excelBook = excel.Workbooks.Open(@path, 0, false);//打开一个工作簿

            Excel.Worksheet sheet1;//打开工作表

            excel.DisplayAlerts = false;

            int sheetCount = excelBook.Sheets.Count;//一共多少个sheet；

            Console.WriteLine("正在对工作表名为：" + fileName + "删除列操作");

            
                for (int a = 1; a <= sheetCount; a++)
                {

                sheet1 = (Excel.Worksheet)excelBook.Sheets[a];
                

                string sheetName = sheet1.Name;
                if (!sheetName.Contains("heet")) {
                    //删除空列
                    deleteEmptyColumn(sheet1);
                    int rowCount = sheet1.UsedRange.Rows.Count;//一共几行
                    //将体格检查文件中的文件路径内容的avi全部替换为mp4
                    replaceAviWithMp4(sheet1, rowCount, fileName);
                    //生命体征名称为空，复制便签的数据  A列是生命体征的，把I列的值复制给C列
                    copyOfTheNote(sheet1, rowCount, fileName);
                    //拆分血压
                    separationOfBloodPressure(excel, sheet1, rowCount, fileName);
                    //比对专项体格检查视频名称加入
                    if ("专项训练".Equals(sheet1.Name))
                    {
                        getSpecialPhysicalExaminationVideoAddress(sheet1, rowCount, fileName);
                    }

                    var columnCount = sheet1.UsedRange.Columns.Count;//一共几列

                    string[] headerNames = new string[] { "是否是音频、视频、图片", "有无交互" };

                    for (int hn = 0; hn < headerNames.Length; hn++)
                    {

                        string headerName = (string)headerNames.GetValue(hn);

                        int j = 0;

                        for (int i = 1; i <= columnCount; i++)

                        {
                            var header = (sheet1.Cells[1, i]).Text.ToString().Trim();

                            if (headerName == header)

                            {

                                j = i; //第几列需要删除

                            }

                            //容错，发现表头为空直接退出循环
                            if (header == "")
                            {
                                break;
                            }
                        }
                        if (j != 0)
                        {
                            sheet1.Range[sheet1.Cells[1, j], sheet1.Cells[1, j]].EntireColumn.Delete();

                        }

                    }

                }
                   

                }

                // excel.Visible = true;
                excelBook.Save();

                excelBook.Close(false);

                excelBook = null;

                //退出Excel程序 
                excel.Quit();
                excel = null;

                // 10.调用GC的垃圾收集方法  
                GC.Collect();

                GC.WaitForPendingFinalizers();
            

            Console.WriteLine("工作表名为：" + fileName + "删除列操作已完成");

        }
        /// <summary>
        /// 获得辅助检查基础库并进行比对数据和赋值
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="rowCount">行数</param>

        public static void getSupplementaryExaminationBaseLibrary(Excel.Worksheet sheet, int rowCount, string fileName) {
            Excel.Application excel_fz = new Excel.Application();//引用Excel对象

            Excel.Workbook excelBook_fz = excel_fz.Workbooks.Open(@"d:\\辅助检查库.xlsx", 0, false);//打开一个工作簿

            Excel.Worksheet sheet_fz;//打开工作表

            excel_fz.DisplayAlerts = false;

            int sheetCount_fz = excelBook_fz.Sheets.Count;//一共多少个sheet；

            string sheet_name = sheet.Name;//需要比对的工作表的名字

            int FineClassNameColumnNumbers = GetColumnNumbers(sheet, "细类名称"); 

            int clinicalSignificanceColumnNumbers = GetColumnNumbers(sheet, "临床意义(用红色字体表示补充项目)");

            for (int a=1;a<= sheetCount_fz; a++) {

                sheet_fz = (Excel.Worksheet)excelBook_fz.Sheets[a];

                
                //待处理工作表名字包含模板名字，即可进行比对逻辑

                if (sheet_name.Contains(sheet_fz.Name)) {

                    Console.WriteLine("正在处理文件名称为“"+ fileName + "”的工作表名为“"+ sheet_name +"”");

                    int rowCount_fz = sheet_fz.UsedRange.Rows.Count;

                    Excel.Range rng1_fz = sheet_fz.Cells.get_Range("B2", "B" + rowCount_fz);

                    object[,] arryItem1_fz = (object[,])rng1_fz.Value2;

                    int FineClassNameColumnNumbers_fz = GetColumnNumbers(sheet_fz, "细类名称");

                    int clinicalSignificanceColumnNumbers_fz = GetColumnNumbers(sheet_fz, "临床意义(用红色字体表示补充项目)");
                    
                    Excel.Range rng1 = sheet.Cells.get_Range("B2", "B" + rowCount);
                    
                    object[,] arryItem1 = (object[,])rng1.Value2;
                    for (int i=2;i<= rowCount;i++) {
                        if (arryItem1[i - 1, 1] == null)
                            arryItem1[i - 1, 1] = "";
                        

                            for (int j = 2; j <= rowCount_fz; j++)
                            {
                                if (arryItem1_fz[j - 1, 1] == null)
                                    arryItem1_fz[j - 1, 1] = "";
                                if (arryItem1_fz[j - 1, 1] == "")
                                    break;
                                if (arryItem1[i - 1, 1].Equals(arryItem1_fz[j - 1, 1])&& clinicalSignificanceColumnNumbers!=0)
                                {

                                    (sheet.Cells[i, clinicalSignificanceColumnNumbers]).value = (sheet_fz.Cells[j, clinicalSignificanceColumnNumbers_fz]).Text.ToString().Trim();
                                (sheet.Cells[i, clinicalSignificanceColumnNumbers]).Font.ColorIndex = 3;
                                break;
                                }

                            }
                       
                        

                    }

                }

            }
            
            excelBook_fz.Save();

            excelBook_fz.Close(false);

            excelBook_fz = null;

            //退出Excel程序 
            excel_fz.Quit();
            excel_fz = null;

            // 10.调用GC的垃圾收集方法  
            GC.Collect();

            GC.WaitForPendingFinalizers();



        }
        /// <summary>
        /// 
        /// 辅助检查使用基础库 的临床意义
        /// </summary>
        /// <param name="path"></param>
        /// <param name="fileName"></param>
        public static void supplementaryExaminationClinicalSignificance(string path, string fileName) {
            Excel.Application excel = new Excel.Application();//引用Excel对象

            Excel.Workbook excelBook = excel.Workbooks.Open(@path, 0, false);//打开一个工作簿

            Excel.Worksheet sheet1;//打开工作表

            excel.DisplayAlerts = false;

            int sheetCount = excelBook.Sheets.Count;//一共多少个sheet；

            for (int a = 1; a <= sheetCount; a++) {
                sheet1 = (Excel.Worksheet)excelBook.Sheets[a];

                //删除空列
                deleteEmptyColumn(sheet1);

                int rowCount = sheet1.UsedRange.Rows.Count;//一共几列

                getSupplementaryExaminationBaseLibrary(sheet1,rowCount,fileName);
            }

            // excel.Visible = true;
            excelBook.Save();

            excelBook.Close(false);

            excelBook = null;

            //退出Excel程序 
            excel.Quit();
            excel = null;

            // 10.调用GC的垃圾收集方法  
            GC.Collect();

            GC.WaitForPendingFinalizers();
            Console.WriteLine("工作表名为：" + fileName + "的表头手术及操作改为操作，运动处方和饮食处方改为体位/活动,膳食已修改完毕");
        }
        /// <summary>
        /// 手术及操作改为操作，运动处方和饮食处方改名体位/活动,膳食，
        /// </summary>
        /// <param name="path"></param>
        public static void modifyThePrescriptionDietPrescription(string path, string fileName)
        {
            Excel.Application excel = new Excel.Application();//引用Excel对象

            Excel.Workbook excelBook = excel.Workbooks.Open(@path, 0, false);//打开一个工作簿

            Excel.Worksheet sheet1;//打开工作表

            excel.DisplayAlerts = false;

            int sheetCount = excelBook.Sheets.Count;//一共多少个sheet；

            Console.WriteLine("正在对工作表名为：" + fileName + "的表头手术及操作改为操作，运动处方和饮食处方改为体位/活动,膳食");
            for (int a = 1; a <= sheetCount; a++)
            {
                sheet1 = (Excel.Worksheet)excelBook.Sheets[a];

                string sheetName = sheet1.Name;

                if (sheetName.Equals("门诊治疗") || sheetName.Equals("临时处置"))
                {

                    var columnCount = sheet1.UsedRange.Columns.Count;//一共几列

                    for (int i = 1; i <= columnCount; i++)

                    {
                        var header = (sheet1.Cells[1, i]).Text.ToString().Trim();
                        if ("手术及操作".Equals(header))
                        {

                            sheet1.Cells[1, i].value = "操作";
                            (sheet1.Cells[1, i]).Font.ColorIndex = 3;
                        }
                        if ("运动处方".Equals(header))
                        {

                            sheet1.Cells[1, i].value = "体位/活动";
                            (sheet1.Cells[1, i]).Font.ColorIndex = 3;
                        }
                        if ("饮食处方".Equals(header))
                        {

                            sheet1.Cells[1, i].value = "膳食";
                            (sheet1.Cells[1, i]).Font.ColorIndex = 3;
                        }

                    }

                }

            }




            // excel.Visible = true;
            excelBook.Save();

            excelBook.Close(false);

            excelBook = null;

            //退出Excel程序 
            excel.Quit();
            excel = null;

            // 10.调用GC的垃圾收集方法  
            GC.Collect();

            GC.WaitForPendingFinalizers();
            Console.WriteLine("工作表名为：" + fileName + "的表头手术及操作改为操作，运动处方和饮食处方改为体位/活动,膳食已修改完毕");
        }
        /// <summary>
        /// 查找sheet标签页
        /// </summary>

        private static Excel.Worksheet FindSheet(Excel.Workbook workbook, string name)
        {
            string[] strArray = name.Split(',');

            for (int i = 1; i <= workbook.Sheets.Count; i++)
            {
                Excel.Worksheet sheet = workbook.Sheets[i];

                if (sheet.Name == name)

                {

                    return sheet;
                }

            }

            return null; 

        }





        /// <summary>  
        /// 获得指定目录及其子目录的所有文件  
        /// </summary>  
        /// <param name="path"></param>  
        /// <returns></returns>  
        public static List<FileInfo> GetAllFilesByDir(string path)
        {
            DirectoryInfo dir = new DirectoryInfo(path);

            //找到该目录下的文件  
            FileInfo[] fi = dir.GetFiles();

            //把FileInfo[]数组转换为List  
            List<FileInfo> list = fi.ToList<FileInfo>();

            //找到该目录下的所有目录里的文件  
            DirectoryInfo[] subDir = dir.GetDirectories();
            foreach (DirectoryInfo d in subDir)
            {
                List<FileInfo> subList = GetFilesByDir(d.FullName);

                foreach (FileInfo subFile in subList)

                {

                    list.Add(subFile);

                }

            }

            return list;

        }

        /// <summary>  
        /// 获得指定目录下的所有文件  
        /// </summary>  
        /// <param name="path"></param>  
        /// <returns></returns>  
        public static List<FileInfo> GetFilesByDir(string path)
        {
            DirectoryInfo di = new DirectoryInfo(path);

            //找到该目录下的文件  
            FileInfo[] fi = di.GetFiles();

            //把FileInfo[]数组转换为List  
            List<FileInfo> list = fi.ToList<FileInfo>();
            return list;
        }
        /// <summary>  
        /// 列出指定目录下及所其有子目录及子目录里更深层目录里的文件（需要递归）  
        /// </summary>  
        /// <param name="path"></param>  
        public static void GetAllFiles(string path)
        {
            DirectoryInfo dir = new DirectoryInfo(path);

            //找到该目录下的文件  
            FileInfo[] fi = dir.GetFiles();
            foreach (FileInfo f in fi)
            {
                Console.WriteLine("完整路径：" + f.FullName.ToString() + " 文件名：" + f.Name );

                //此处增加容错，需求方提供的文件夹中存在包含以下符号的文件名，过滤掉
                if (!f.FullName.ToString().Contains("$")) {

                 if (f.Name.Contains("体格检查") && f.FullName.ToString().Contains("xls"))

                {

                    openExcelAndDeleteColumns(f.FullName.ToString(), f.Name);

                }
                if (f.Name.Contains("治疗") && f.FullName.ToString().Contains("xls"))

                {
                    //手术及操作改为操作，运动处方和饮食处方改名体位/活动,膳食
                    modifyThePrescriptionDietPrescription(f.FullName.ToString(), f.Name);

                }
               if (f.Name.Contains("辅助检查.xls"))

                {
                    //辅助检查使用基础库的临床意义
                    supplementaryExaminationClinicalSignificance(f.FullName.ToString(), f.Name);

                }

                }


            }

            //找到该目录下的所有目录再递归 
            DirectoryInfo[] subDir = dir.GetDirectories();

            foreach (DirectoryInfo d in subDir)

            {

                GetAllFiles(d.FullName);

            }
        }
        /// <summary>
        /// 将列名字转换成列序号
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columnNames"></param>
        /// <returns></returns>
        public static int GetColumnNumbers(Excel.Worksheet sheet, string columnNames)
        {
            int columnNumbers = 0;

            var columnCount = sheet.UsedRange.Columns.Count;
            for (int i = 1; i <= columnCount; i++)
            {
                var header = (sheet.Cells[1, i]).Text.ToString().Trim();
                if (columnNames.Equals(header))
                {
                    columnNumbers = i;
                    break;
                }
            }

            return columnNumbers;
        }
        /// <summary>
        /// 删除空列
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static void deleteEmptyColumn(Excel.Worksheet sheet)
        {


            //Excel.Worksheet sheet1;
            Console.WriteLine("正在删除工作表"+ sheet.Name+ "中的空列");
            var columnCount = sheet.UsedRange.Columns.Count;//一共几列

            int beginIndex = 0, endIndex = 0;
            for (int j = columnCount; j >= 1; j--)
            {

                if (!string.IsNullOrWhiteSpace((sheet.Cells[1, j]).Text.ToString()))
                {

                    break;

                }

                if (endIndex == 0)
                {
                    endIndex = j;
                }
                beginIndex = j;

            }
            if (beginIndex != 0 && endIndex != 0)
            {
                sheet.Range[(sheet.Cells[1, beginIndex]), (sheet.Cells[1, endIndex])].EntireColumn.Delete(1);
            }
            Console.WriteLine("工作表" + sheet.Name + "中的空列已经删除完毕");
        }

    }
}
