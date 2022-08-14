using System;
using System.Runtime.InteropServices;
using System.IO;
using System.Collections;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;

namespace Subassembly
{
    class Program
    {
        #region MyRegion
        public static string title = "待开发";
        public static string subject = "待开发";
        public static int count = 1;
        public static string material = " ";
        public static string docNum = "待开发";
        public static string massName = "待开发";
        public static string quality = " ";
        public static string myType = "待开发";
        public static string revNum = "待开发";
        public static string fileName = "待开发";
        public static string qtemp = " ";
        //public static string myStemName;
        public static bool flag = false;
        public static bool dflag = true;
        public static bool tdflag = true;            //temp
        public static bool isSubassembly;
        public static string mypath;
        /*public static ArrayList list = new ArrayList();    */                             // 定义一个用于存放数据的列表
        public static ArrayList templist = new ArrayList();
        public static DesignManager.Application dmapplication = new DesignManager.Application();
        public static DesignManager.Document document = null;
        public static DesignManager.PropertySets propertySets = null;
        //part
        public static SolidEdgePart.PartDocument partDocument = null;
        public static SolidEdgePart.Models models = null;
        public static SolidEdgePart.Model model = null;
        public static double density = 0;
        public static double accuracy = 0;
        public static double volume = 0;
        public static double area = 0;
        public static double mass = 0;
        public static Array cetnerOfGravity = Array.CreateInstance(typeof(double), 3);
        public static Array centerOfVolumne = Array.CreateInstance(typeof(double), 3);
        public static Array globalMomentsOfInteria = Array.CreateInstance(typeof(double), 6);     // Ixx, Iyy, Izz, Ixy, Ixz and Iyz 
        public static Array principalMomentsOfInteria = Array.CreateInstance(typeof(double), 3);  // Ixx, Iyy and Izz
        public static Array principalAxes = Array.CreateInstance(typeof(double), 9);              // 3 axes x 3 coords
        public static Array radiiOfGyration = Array.CreateInstance(typeof(double), 9);            // 3 axes x 3 coords
        public static double relativeAccuracyAchieved = 0;
        public static int status = 0;
        //assem
        public static Array centerOfMass = Array.CreateInstance(typeof(double), 3);
        public static Array centerOfVolume = Array.CreateInstance(typeof(double), 3);
        public static Array globalMoments = Array.CreateInstance(typeof(double), 3);
        public static Array principalAxis1 = Array.CreateInstance(typeof(double), 3);
        public static Array principalAxis2 = Array.CreateInstance(typeof(double), 3);
        public static Array principalAxis3 = Array.CreateInstance(typeof(double), 3);
        public static Array principalMoments = Array.CreateInstance(typeof(double), 3);
        public static bool isSick;
        public static bool updateStatus;

        //结构
        public struct Infolist
        {
            public string temptitle;
            public string tempsubject;
            public int tempcount;
            public string tempdocNum;
            public string tempmyType;
            public string temprevNum;
            public string tempmaterial;
            public string tempquality;
        }
        //
        public static SolidEdgeFramework.Variables variables = null;
        public static SolidEdgeFramework.variable variable = null;

        #endregion

        [STAThread]
        static void Main(string[] args)
        {
            Console.Title = "SolidEdge 技术交流群：377600942 提供";
            //Console.Title = " ";
            SolidEdgeFramework.Application application = null;
            SolidEdgeAssembly.AssemblyDocument assemblyDocument = null;
            Console.WriteLine(" \n");
            Console.WriteLine("生成的明细表存放于桌面的temp明细表文件夹中\n");
            Console.WriteLine("如果程序执行到一半被人为叉掉的话，需要去后台手动把设计管理器关掉\n");
            Console.WriteLine("  \n");
            Console.WriteLine("\t\t\t\t\t\t明细表\n");
            Console.WriteLine("标题\t\t主题\t\t数量\t材料\t文档号\t\t质量\t\t类别\t版本\t文件名\n");
            try
            {
                //list.Clear();
                application = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
                assemblyDocument = application.ActiveDocument as SolidEdgeAssembly.AssemblyDocument;

                application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)33091);
                //application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)45077);

                mypath = application.UserName;

                //createfile
                if (! System.IO.Directory.Exists("C:\\Users\\" + mypath + "\\Desktop\\temp明细表"))
                {
                    System.IO.Directory.CreateDirectory("C:\\Users\\" + mypath + "\\Desktop\\temp明细表");
                }

                title = assemblyDocument.SummaryInfo.Title;
                subject = assemblyDocument.SummaryInfo.Subject;
                myType = assemblyDocument.SummaryInfo.Category;
                docNum = assemblyDocument.SummaryInfo.DocumentNumber;
                revNum = assemblyDocument.SummaryInfo.RevisionNumber;
                fileName = assemblyDocument.DisplayName;

                mypath = "C:\\Users\\" + mypath + "\\Desktop\\temp明细表\\" + subject + "-" + docNum + "-" + revNum + ".xlsx";
                File.AppendAllText(mypath, mypath);
                File.Delete(mypath);
                qtemp = docNum + "-" + revNum;

                assemblyDocument.PhysicalProperties.GetAssemblyPhysicalProperties(
                                 Mass: out mass,
                                 Volume: out volume,
                                 Area: out area,
                                 CenterOfMass: ref centerOfMass,
                                 CenterOfVolume: ref centerOfVolume,
                                 GlobalMoments: ref globalMoments,
                                 PrincipalAxis1: ref principalAxis1,
                                 PrincipalAxis2: ref principalAxis2,
                                 PrincipalAxis3: ref principalAxis3,
                                 PrincipalMoments: ref principalMoments,
                                 RadiiOfGyration: ref radiiOfGyration,
                                 IsSick: out isSick,
                                 UpdateStatus: out updateStatus);
                quality = mass.ToString("F3") + " kg";

                Console.WriteLine("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t\t{6}\t{7}\t{8}\n", title, subject, count, material, docNum, quality, myType, revNum, fileName);

                flag = true;

                Infolist p1;
                p1.temptitle = title;
                p1.tempsubject = subject;
                p1.tempcount = 1;
                p1.tempdocNum = docNum;
                p1.tempmyType = myType;
                p1.temprevNum = revNum;
                p1.tempmaterial = material;
                p1.tempquality = quality;

                templist.Add(p1);

                DisplayOn(assemblyDocument);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.WriteLine("未打开装配模型或solidedge活动界面非装配模型！！！\n");
                Console.WriteLine("\n");
            }
            //写入Excel文件
            Write_data();

            Console.WriteLine();
            Console.WriteLine("按任意键退出!!!");
            dmapplication.Quit();
            templist.Clear();
            Console.ReadKey();
        }


 
        static void Write_data(string mxb)  
        {
            if (flag)
            {
                File.AppendAllText(mypath, "标题,主题,数量,材料,文档号,质量,类别,版本,文件名\r\n");
                flag = false;
            }
            File.AppendAllText(mypath, mxb);
        }

        static void Write_data()
        {
            IWorkbook wkbook = new HSSFWorkbook();
            ISheet sheet = wkbook.CreateSheet(qtemp);
            sheet.SetColumnWidth(0, 15 * 320);
            sheet.SetColumnWidth(1, 15 * 300);
            sheet.SetColumnWidth(2, 15 * 120);
            sheet.SetColumnWidth(3, 15 * 260);
            sheet.SetColumnWidth(4, 15 * 260);
            sheet.SetColumnWidth(5, 15 * 185);
            sheet.SetColumnWidth(6, 15 * 120);
            bool wt = true;
            Infolist tempdata;
            Console.WriteLine(templist.Count);

            for (int i = 0; i < templist.Count + 1; i++)
            {
                IRow row = sheet.CreateRow(i);
                if (wt)
                {
                    row.CreateCell(0).SetCellValue("代号");
                    row.CreateCell(1).SetCellValue("名称");
                    row.CreateCell(2).SetCellValue("数量");
                    row.CreateCell(3).SetCellValue("材料");
                    row.CreateCell(4).SetCellValue("编码");
                    row.CreateCell(5).SetCellValue("重量");
                    row.CreateCell(6).SetCellValue("备注");
                    row.CreateCell(7).SetCellValue("版本");
                    wt = false;
                }
                else
                {
                    //写入数据
                    //row.CreateCell(0).SetCellValue("标题");
                    tempdata = (Infolist)templist[i - 1];
                    row.CreateCell(0).SetCellValue((string)tempdata.temptitle);
                    row.CreateCell(1).SetCellValue((string)tempdata.tempsubject);
                    row.CreateCell(2).SetCellValue((int)tempdata.tempcount);
                    row.CreateCell(3).SetCellValue((string)tempdata.tempmaterial);
                    row.CreateCell(4).SetCellValue((string)tempdata.tempdocNum);
                    row.CreateCell(5).SetCellValue((string)tempdata.tempquality);
                    row.CreateCell(6).SetCellValue((string)tempdata.tempmyType);
                    row.CreateCell(7).SetCellValue((string)tempdata.temprevNum);
                  
                }              
            }
            //将信息写入文件
            using (FileStream fsWrite = File.OpenWrite(mypath))
            {
                wkbook.Write(fsWrite);
            }
        }

        //测试函数
        static void DisplayOn(SolidEdgeAssembly.AssemblyDocument assemblyDocument)
        {
            Infolist p2, p3;
            for (int i = 1; i <= assemblyDocument.Occurrences.Count; i++)
            {
                propertySets = (DesignManager.PropertySets)dmapplication.PropertySets;
                propertySets.Open(assemblyDocument.Occurrences.Item(i).PartFileName, true);
                title = assemblyDocument.Occurrences.Item(i).OccurrenceDocument.SummaryInfo.Title;
                subject = assemblyDocument.Occurrences.Item(i).OccurrenceDocument.SummaryInfo.Subject;
                myType = assemblyDocument.Occurrences.Item(i).OccurrenceDocument.SummaryInfo.Category;
                docNum = assemblyDocument.Occurrences.Item(i).OccurrenceDocument.SummaryInfo.DocumentNumber;
                revNum = assemblyDocument.Occurrences.Item(i).OccurrenceDocument.SummaryInfo.RevisionNumber;
                fileName = assemblyDocument.Occurrences.Item(i).Name;
                fileName = fileName.Substring(0, fileName.IndexOf(":"));

                partDocument = assemblyDocument.Occurrences.Item(i).PartDocument as SolidEdgePart.PartDocument;

                isSubassembly = assemblyDocument.Occurrences.Item(i).Subassembly;
                if (!isSubassembly)
                {
                    material = propertySets.Item["MechanicalModeling"].Item["Material"].value;

                    try
                    {

                        if (partDocument != null)
                        {
                            models = partDocument.Models;
                            model = models.Item(1);

                            model.GetPhysicalProperties(
                                 Status: out status,
                                 Density: out density,
                                 Accuracy: out accuracy,
                                 Volume: out volume,
                                 Area: out area,
                                 Mass: out mass,
                                 CenterOfGravity: ref cetnerOfGravity,
                                 CenterOfVolume: ref centerOfVolumne,
                                 GlobalMomentsOfInteria: ref globalMomentsOfInteria,
                                 PrincipalMomentsOfInteria: ref principalMomentsOfInteria,
                                 PrincipalAxes: ref principalAxes,
                                 RadiiOfGyration: ref radiiOfGyration,
                                 RelativeAccuracyAchieved: out relativeAccuracyAchieved);
                        }
                        quality = mass.ToString("F3") + " kg";
                    }
                    catch (Exception)
                    {
                        quality = "//";
                    }
                }

                p2.temptitle = title;
                p2.tempsubject = subject;
                p2.tempcount = 1;
                p2.tempdocNum = docNum;
                p2.tempmyType = myType;
                p2.temprevNum = revNum;
                p2.tempmaterial = material;
                p2.tempquality = quality;

                if (isSubassembly == true)
                {
                    assemblyDocument.Occurrences.Item(i).OccurrenceDocument.PhysicalProperties.GetAssemblyPhysicalProperties(
                                 Mass: out mass,
                                 Volume: out volume,
                                 Area: out area,
                                 CenterOfMass: ref centerOfMass,
                                 CenterOfVolume: ref centerOfVolume,
                                 GlobalMoments: ref globalMoments,
                                 PrincipalAxis1: ref principalAxis1,
                                 PrincipalAxis2: ref principalAxis2,
                                 PrincipalAxis3: ref principalAxis3,
                                 PrincipalMoments: ref principalMoments,
                                 RadiiOfGyration: ref radiiOfGyration,
                                 IsSick: out isSick,
                                 UpdateStatus: out updateStatus);
                    quality = mass.ToString("F3") + " kg";

                    p2.temptitle = title;
                    p2.tempsubject = subject;
                    p2.tempcount = 1;
                    p2.tempdocNum = docNum;
                    p2.tempmyType = myType;
                    p2.temprevNum = revNum;
                    p2.tempmaterial = " ";
                    p2.tempquality = quality;

                    for (int b = 0; b < templist.Count; b++)
                    {
                        p3 = (Infolist)templist[b];
                        if (p3.tempdocNum.Contains(docNum))
                        {
                            templist.Remove(templist[b]);
                            p3.tempcount = p3.tempcount + 1;
                            templist.Insert(b, p3);
                            dflag = false;
                            break;
                        }
                    }
                    if (dflag)
                    {
                        templist.Add(p2);
                    }
                    if (!(docNum.StartsWith("H") || docNum.StartsWith("81")))
                    {
                        dflag = true;
                        DisplayOn(assemblyDocument.Occurrences.Item(i).OccurrenceDocument);
                    }
                }
                else
                {
                    for (int b = 0; b < templist.Count; b++)
                    {
                        p3 = (Infolist)templist[b];
                        if (p3.tempdocNum.Contains(docNum))
                        {
                            templist.Remove(templist[b]);
                            p3.tempcount = p3.tempcount + 1;
                            templist.Insert(b, p3);
                            dflag = false;
                            break;
                        }
                    }
                    if (dflag)
                    {
                        templist.Add(p2);
                    }
                }
                dflag = true;
            }
        }


    }
}



