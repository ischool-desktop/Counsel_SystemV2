using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using Aspose.Cells;
using System.IO;
using FISCA.Presentation.Controls;
using System.Windows.Forms;
using FISCA.Data;
using System.Data;



namespace IdentityReport_whsh
{
    class SingleParent
    {
        List<string> _StudentIDList;
        BackgroundWorker _bgLoadData;  

        public SingleParent(List<string> StudentIDList)
        {
            _bgLoadData = new BackgroundWorker();
            _bgLoadData.DoWork += _bgLoadData_DoWork;
            _bgLoadData.ProgressChanged += _bgLoadData_ProgressChanged;
            _bgLoadData.WorkerReportsProgress = true;
            _bgLoadData.RunWorkerCompleted += _bgLoadData_RunWorkerCompleted;

            //學生編號
            _StudentIDList = StudentIDList;
            //載入資料
            _bgLoadData.RunWorkerAsync();
        }

        void _bgLoadData_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("單親家庭統計表", e.ProgressPercentage);
        }

        void _bgLoadData_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Workbook wb = (Workbook)e.Result;

            SaveFileDialog sd = new SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = "單親家庭統計表.xlsx";
            sd.Filter = "Excel檔案 (*.xlsx)|*.xlsx|所有檔案 (*.*)|*.*";
            if (sd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    wb.Save(sd.FileName);
                    System.Diagnostics.Process.Start(sd.FileName);
                }
                catch
                {
                    MessageBox.Show("指定路徑無法存取。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            FISCA.Presentation.MotherForm.SetStatusBarMessage("單親家庭統計表 產生完成");                        
        }

        void _bgLoadData_DoWork(object sender, DoWorkEventArgs e)
        {
            List<K12.Data.StudentRecord> stu_list = K12.Data.Student.SelectByIDs(_StudentIDList);

            // 找K12.Data 資料 舊版本寫法
            //List<K12.Data.ParentRecord> pr_list = K12.Data.Parent.SelectByStudentIDs(_StudentIDList);

            // 從輔導綜合紀錄表 UDT $ischool.counsel.yearly_data 的項目 (key =家庭狀況_父母關係) 撈單親資料
            string sql = "select *from $ischool.counsel.yearly_data where key ='家庭狀況_父母關係' and ref_student_id in ('" + string.Join("','", _StudentIDList.ToArray()) + "')";

            QueryHelper qh = new QueryHelper();

            DataTable dt = qh.Select(sql);

            Workbook wb = new Workbook(new MemoryStream(Properties.Resources.單親家庭統計表樣板));

            Worksheet ws = wb.Worksheets["單親家庭統計表"];

            Cells cells = ws.Cells;
            
            int row_counter = 1;
            int success_count = 0;

            foreach (K12.Data.StudentRecord sr in stu_list)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    if ("" + dr["ref_student_id"] == sr.ID)
                    {
                        if (sr.Class != null)
                        {
                            if (sr.Class.GradeYear != null)
                            {
                                if ("" + dr["g" + sr.Class.GradeYear] == "離婚" || "" + dr["g" + sr.Class.GradeYear] == "單親-父" || "" + dr["g" + sr.Class.GradeYear] == "單親-母" || "" + dr["g" + sr.Class.GradeYear] == "單親-其他親屬")
                                {
                                    cells[row_counter, 0].Value = sr.StudentNumber;
                                    cells[row_counter, 1].Value = sr.Name;
                                    cells[row_counter, 2].Value = sr.Gender;
                                    cells[row_counter, 3].Value = sr.Class != null? sr.Class.Name:"";
                                    cells[row_counter, 4].Value = sr.SeatNo;
                                    cells[row_counter, 5].Value = "" + dr["g" + sr.Class.GradeYear];

                                    row_counter++;

                                    if (row_counter > 1)
                                    {
                                        Style sy = cells[1, 0].GetStyle();

                                        cells[row_counter, 0].SetStyle(sy);
                                        cells[row_counter, 1].SetStyle(sy);
                                        cells[row_counter, 2].SetStyle(sy);
                                        cells[row_counter, 3].SetStyle(sy);
                                        cells[row_counter, 4].SetStyle(sy);
                                        cells[row_counter, 5].SetStyle(sy);                                        
                                    }
                                }
                            }
                        }
                    }
                }

                try
                {
                    success_count++;

                    int progress = (success_count * 100 / dt.Rows.Count);

                    _bgLoadData.ReportProgress(progress);
                }
                catch (Exception ex)
                {
                    MsgBox.Show(ex.Message, "錯誤", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                }
            }

            #region 找K12.Data 資料 舊版本寫法 
            //foreach (K12.Data.ParentRecord pr in pr_list)
            //{
            //    //當父歿、母歿 判定為單親名單
            //    if (pr.Father.Living == "歿" || pr.Mother.Living == "歿")
            //    {
            //        cells[row_counter, 0].Value = pr.Student.StudentNumber;
            //        cells[row_counter, 1].Value = pr.Student.Name;
            //        cells[row_counter, 2].Value = pr.Student.Gender;
            //        cells[row_counter, 3].Value = pr.Student.Class.Name;
            //        cells[row_counter, 4].Value = pr.Student.SeatNo;
            //        cells[row_counter, 5].Value = pr.Father.Living == "歿" ? "父歿" : "母歿";


            //        row_counter++;

            //        if (row_counter > 1)
            //        {
            //            Style sy = cells[1, 0].GetStyle();

            //            cells[row_counter, 0].SetStyle(sy);
            //            cells[row_counter, 1].SetStyle(sy);
            //            cells[row_counter, 2].SetStyle(sy);
            //            cells[row_counter, 3].SetStyle(sy);
            //            cells[row_counter, 4].SetStyle(sy);
            //            cells[row_counter, 5].SetStyle(sy);
            //            cells[row_counter, 6].SetStyle(sy);
            //            cells[row_counter, 7].SetStyle(sy);
            //            cells[row_counter, 8].SetStyle(sy);
            //        }

            //    }

            //    try
            //    {
            //        success_count++;

            //        int progress = (success_count * 100 / pr_list.Count);

            //        _bgLoadData.ReportProgress(progress);
            //    }
            //    catch (Exception ex)
            //    {
            //        MsgBox.Show(ex.Message, "錯誤", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
            //    }
            //} 
            #endregion

            e.Result = wb;
        }
    }
}

