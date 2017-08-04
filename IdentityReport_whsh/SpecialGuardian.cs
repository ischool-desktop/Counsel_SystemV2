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

namespace IdentityReport_whsh
{
    class SpecialGuardian
    {
        List<string> _StudentIDList;
        BackgroundWorker _bgLoadData;  

        public SpecialGuardian(List<string> StudentIDList)
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
            FISCA.Presentation.MotherForm.SetStatusBarMessage("特殊監護名冊", e.ProgressPercentage);
        }

        void _bgLoadData_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Workbook wb = (Workbook)e.Result;

            SaveFileDialog sd = new SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = "特殊監護名冊.xlsx";
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
            FISCA.Presentation.MotherForm.SetStatusBarMessage("特殊監護名冊 產生完成");                        
        }

        void _bgLoadData_DoWork(object sender, DoWorkEventArgs e)
        {
            //List<K12.Data.StudentRecord> sr_list = K12.Data.Student.SelectByIDs(_StudentIDList);

            

            List<K12.Data.ParentRecord> pr_list = K12.Data.Parent.SelectByStudentIDs(_StudentIDList);

            K12.Data.ParentRecord pr_compare = new K12.Data.ParentRecord();


            
            Workbook wb = new Workbook(new MemoryStream(Properties.Resources.特殊監護名冊樣板));

            Worksheet ws = wb.Worksheets["特殊監護名冊"];

            Cells cells = ws.Cells;
            
            int row_counter = 2;
            int success_count = 0;

            foreach (K12.Data.ParentRecord pr in pr_list)
            {

                //當監護人 身分字號 與父、母 身分字號皆不同時，判定為特殊監護
                if (pr.Custodian.IDNumber != pr.Father.IDNumber && pr.Custodian.IDNumber != pr.Mother.IDNumber)
                {
                    cells[row_counter, 0].Value = pr.Student.StudentNumber;
                    cells[row_counter, 1].Value = pr.Student.Name;
                    cells[row_counter, 2].Value = pr.Student.Gender;
                    cells[row_counter, 3].Value = pr.Student.Class.Name;
                    cells[row_counter, 4].Value = pr.Student.SeatNo;
                    cells[row_counter, 5].Value = pr.Father.Name;
                    cells[row_counter, 6].Value = pr.Mother.Name;
                    cells[row_counter, 7].Value = pr.Custodian.Name;
                    cells[row_counter, 8].Value = pr.Custodian.Relationship;

                    row_counter++;

                    if (row_counter > 2)
                    {
                        Style sy = cells[2,0].GetStyle();

                        cells[row_counter, 0].SetStyle(sy);
                        cells[row_counter, 1].SetStyle(sy);
                        cells[row_counter, 2].SetStyle(sy);
                        cells[row_counter, 3].SetStyle(sy);
                        cells[row_counter, 4].SetStyle(sy);
                        cells[row_counter, 5].SetStyle(sy);
                        cells[row_counter, 6].SetStyle(sy);
                        cells[row_counter, 7].SetStyle(sy);
                        cells[row_counter, 8].SetStyle(sy);
                    }

                    try
                    {
                        success_count++;

                        int progress = (success_count * 100 / pr_list.Count);

                        _bgLoadData.ReportProgress(progress);
                    }
                    catch (Exception ex)
                    {
                        MsgBox.Show(ex.Message, "錯誤", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                    }
                }
            }

            e.Result = wb;
        }

        
        }
}

