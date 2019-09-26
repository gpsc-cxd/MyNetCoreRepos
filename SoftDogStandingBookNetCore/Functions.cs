/////////////////////////////////////////////////////////////////////////
//// All rights reserved.
//// author: Adminstrator
//// File: Functions.cs
//// Summary: Functions
//// Date: 2019/9/19 13:37:59
//////////////////////////////////////////////////////////////////////////
using Aspose.Words;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace SoftDogStandingBookNetCore
{
    public class Functions
    {
        public static string modelDoc = AppDomain.CurrentDomain.BaseDirectory + "Model\\Model.doc";
        public static string modelExcel = AppDomain.CurrentDomain.BaseDirectory + "Model\\Model.xlsx";

        public static string ChangeDateString(string date)
        {
            var sp = date.Split('/');
            string result = string.Format("{0}/{1}/{2}", sp[0], Convert.ToInt32(sp[1]).ToString("00"), Convert.ToInt32(sp[2]).ToString("00"));
            return result;
        }
        public static void ExportWord(string fileName, Datas datas)
        {
            File.Copy(modelDoc, fileName, true);
            //MemoryStream ms = new MemoryStream();
            //FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            //byte[] bytes = new byte[fs.Length];
            //fs.Read(bytes, 0, (int)fs.Length);
            //ms.Write(bytes, 0, (int)fs.Length);
            //fs.Close();
            Document doc = new Document(fileName);
            DocumentBuilder db = new DocumentBuilder(doc);
            WriteDoc(doc, db, datas);
            //软件锁编号
            db.MoveToBookmark("dogcode");
            var run = new Run(doc, datas.Dogcode);
            run.Font.Size = 14;
            run.Font.Color = System.Drawing.Color.Blue;
            db.InsertNode(run);
            //备注
            db.MoveToBookmark("remark");
            string remark = string.Empty;
            if (datas.Servicetype.Contains("新领取"))
            {
                remark = string.Format("新申请1个{0}", datas.Dogtype);
            }
            else
            {
                remark = datas.Dogtype + datas.Servicetype;
            }
            db.InsertNode(new Run(doc, remark));
            //保存
            doc.Save(fileName, Aspose.Words.SaveFormat.Doc);
        }
        public static void ExportWord(string outputPath, IEnumerable<Datas> datases)
        {
            //先确定文件数
            //根据测区来分文件
            var files = datases.GroupBy(a => a.Regionalcode);
            var datetime = DateTime.Now;
            string date = datetime.Year.ToString() + datetime.Month.ToString("00") + datetime.Day.ToString("00");
            foreach (var f in files)
            {
                string filename = string.Format("{0}-加密锁授权申请表-{1}-{2}.doc", date, f.ElementAt(0).Compname, f.ElementAt(0).Regionalname);
                var destFile = Path.Combine(outputPath, filename);
                File.Copy(modelDoc, destFile, true);//复制文件
                var datas = f.ElementAt(0);
                Document doc = new Document(destFile);
                DocumentBuilder db = new DocumentBuilder(doc);
                WriteDoc(doc, db, datas);
                //锁编号
                string dogCode = string.Empty;
                foreach (var d in f)
                {
                    dogCode += d.Dogcode + "；";
                }
                dogCode = dogCode.TrimEnd('；');
                db.MoveToBookmark("dogcode");
                var run = new Run(doc, dogCode);
                run.Font.Size = 14;
                run.Font.Color = System.Drawing.Color.Blue;
                db.InsertNode(run);
                //备注
                string remark = string.Empty;
                var dt = f.GroupBy(a => a.Dogtype);
                foreach (var d in dt)
                {
                    var st = d.GroupBy(a => a.Servicetype);
                    foreach (var s in st)
                    {
                        remark += string.Format("{0}{1}{2}个；", d.Key, s.Key, s.Count());
                    }
                }
                remark = remark.TrimEnd('；');
                db.MoveToBookmark("remark");
                db.InsertNode(new Run(doc, remark));
                doc.Save(destFile, Aspose.Words.SaveFormat.Doc);
            }
        }
        private static void WriteDoc(Document doc, DocumentBuilder db, Datas datas)
        {
            db.MoveToBookmark("name");
            db.InsertNode(new Run(doc, datas.Name));
            db.MoveToBookmark("applydate");
            db.InsertNode(new Run(doc, datas.Applydate));
            db.MoveToBookmark("compname");
            db.InsertNode(new Run(doc, datas.Compname));
            db.MoveToBookmark("phonenumber");
            db.InsertNode(new Run(doc, datas.Phonenumber));
            db.MoveToBookmark("softwarename");
            db.InsertNode(new Run(doc, datas.Softwarename));
            //授权期限
            var expdate = datas.Expirationdate.Split('/');
            db.MoveToBookmark("expyear");
            var run = new Run(doc, expdate[0]);
            run.Font.Underline = Underline.Single;
            db.InsertNode(run);
            db.MoveToBookmark("expmonth");
            run = new Run(doc, expdate[1].Length == 1 ? "0" + expdate[1] : expdate[1]);
            run.Font.Underline = Underline.Single;
            db.InsertNode(run);
            db.MoveToBookmark("expday");
            run = new Run(doc, expdate[2].Length == 1 ? "0" + expdate[2] : expdate[2]);
            run.Font.Underline = Underline.Single;
            db.InsertNode(run);
            db.MoveToBookmark("regional");
            run = new Run(doc, string.Format("{0}({1})", datas.Regionalname, datas.Regionalcode));
            run.Font.Underline = Underline.Single;
            db.InsertNode(run);

        }
    }
}
