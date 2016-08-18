using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.Data;
using System.Data;
using Aspose.Cells;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using System.IO;

namespace CounselTools
{
   public class Utility
    {
       /// <summary>
        /// 透過學生ID取得輔導-DataRow
       /// </summary>
       /// <param name="StudentIDList"></param>
       /// <returns></returns>
       public static Dictionary<string, List<DataRow>> GetABCardAnswerDataByStudentIDList(List<string> StudentIDList,string Type)
       {
           Dictionary<string, List<DataRow>> retVal = new Dictionary<string, List<DataRow>>();
           string query = "";
           switch (Type)
           { 
               case "multiple_record":
                   query = "select * from $ischool.counsel.multiple_record";
                   break;

               case "priority_data":
                   query = "select * from $ischool.counsel.priority_data";
                   break;

               case "relative":
                   query = "select * from $ischool.counsel.relative";
                   break;

               case "semester_data":
                   query = "select * from $ischool.counsel.semester_data";
                   break;

               case "sibling":
                   query = "select * from $ischool.counsel.sibling";
                   break;

               case "single_record":
                   query = "select * from $ischool.counsel.single_record";
                   break;

               case "yearly_data":
                   query = "select * from $ischool.counsel.yearly_data";
                   break;
           }

           if (query != "")
           {
               query += " where ref_student_id in("+string.Join(",",StudentIDList.ToArray())+")";

               QueryHelper qh = new QueryHelper();
               DataTable dt = qh.Select(query);

               foreach (DataRow dr in dt.Rows)
               {
                   string id = dr["ref_student_id"].ToString();
                   if (!retVal.ContainsKey(id))
                       retVal.Add(id, new List<DataRow>());

                   retVal[id].Add(dr);               
               }
           }

           return retVal;
       }

       /// <summary>
       /// 匯出 Excel
       /// </summary>
       /// <param name="inputReportName"></param>
       /// <param name="inputXls"></param>
       public static void CompletedXls(string inputReportName, Workbook inputXls)
       {
           string reportName = inputReportName;

           string path = Path.Combine(Application.StartupPath, "Reports");
           if (!Directory.Exists(path))
               Directory.CreateDirectory(path);
           path = Path.Combine(path, reportName + ".xls");

           Workbook wb = inputXls;

           if (File.Exists(path))
           {
               int i = 1;
               while (true)
               {
                   string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                   if (!File.Exists(newPath))
                   {
                       path = newPath;
                       break;
                   }
               }
           }

           try
           {
               wb.Save(path, SaveFormat.Excel97To2003);
               System.Diagnostics.Process.Start(path);
           }
           catch
           {
               SaveFileDialog sd = new SaveFileDialog();
               sd.Title = "另存新檔";
               sd.FileName = reportName + ".xls";
               sd.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
               if (sd.ShowDialog() == DialogResult.OK)
               {
                   try
                   {
                       wb.Save(sd.FileName, SaveFormat.Excel97To2003);

                   }
                   catch
                   {
                       MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                       return;
                   }
               }
           }
       }


       public static List<ClassStudent> GetClassStudentByStudentIDList(List<string> StudentIDList)
       {
           List<ClassStudent> retVal = new List<ClassStudent>();
           if (StudentIDList.Count > 0)
           {
               QueryHelper qh = new QueryHelper();
               string query = "select student.id as sid,class_name,class.grade_year,seat_no,student_number,student.name as sname from student left join class on student.ref_class_id=class.id where student.id in (" + string.Join(",", StudentIDList.ToArray()) + ") order by class.grade_year,class_name,seat_no,student_number";
               DataTable dt = qh.Select(query);
               foreach (DataRow dr in dt.Rows)
               {
                   ClassStudent cs = new ClassStudent();
                   cs.ClassName = dr["class_name"].ToString();
                   cs.SeatNo = dr["seat_no"].ToString();
                   cs.StudentID = dr["sid"].ToString();
                   cs.StudentNumber = dr["student_number"].ToString();
                   cs.StudentName = dr["sname"].ToString();
                   cs.GradeYear=0;
                   cs.GradeYearDisplay = 0;
                   int gY;
                   if (int.TryParse(dr["grade_year"].ToString(), out gY))
                   {
                       cs.GradeYear = gY;
                       cs.GradeYearDisplay = gY;

                       if (gY == 7 || gY == 10)
                           cs.GradeYear = 1;

                       if (gY == 8 || gY == 11)
                           cs.GradeYear = 2;

                       if (gY == 9 || gY == 12)
                           cs.GradeYear = 3;
                   }
                       
                   retVal.Add(cs);
               }
           }
           return retVal;
       }

       
    }
}
