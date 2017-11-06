using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

using FISCA.Presentation.Controls;
using Aspose.Cells;
using K12.Data;


namespace AfterGratuationStatics_whsh.Forms
{
    public partial class Step1_Form : BaseForm
    {
        BackgroundWorker _bw = new BackgroundWorker();

        // 檔案位置
        public string source_data = "";

        public Step1_Form()
        {
            InitializeComponent();
        }

        // 取消
        private void buttonX3_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        //選擇檔案來源
        private void buttonX1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ope = new OpenFileDialog();
            ope.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            if (ope.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }
            else
            {
                textBoxX1.Text = ope.FileName;
            }
        }

        //匯入
        private void buttonX2_Click(object sender, EventArgs e)
        {
            source_data = textBoxX1.Text;

            // 錯誤資料List
            List<string> errorList = new List<string>();

            // 若沒選取來源檔案，中止程序
            if (source_data == "")
            {
                MsgBox.Show("請選擇來源檔案", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 載入匯入檔案
            Workbook wb = new Workbook(source_data);

            // 文件打開舊方法，已過時， 現在統一使用上面的new Workbook(textBoxX1.Text) 的方式，另外 請使用最新 Aspose.Cell_201402 ，可以避免很多讀取錯誤Bug
            //wb.Open(textBoxX1.Text, FileFormatType.Excel2007Xlsx);

            #region 資料驗證

            #region 驗證畢業生榜單
            // 檢查樣版格式是否正確
            if (wb.Worksheets["畢業生榜單"] == null)
            {
                MsgBox.Show("Excel檔案 沒有'畢業生榜單' 頁籤，請確認樣板格式正確", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Worksheet ws = wb.Worksheets["畢業生榜單"];

            Cells cells = ws.Cells;

            if ("" + cells[0, 0].Value != "學號")
            {
                errorList.Add("頁籤:畢業生榜單，第一行表頭應為:學號");
            }
            if ("" + cells[0, 1].Value != "科別")
            {
                errorList.Add("頁籤:畢業生榜單，第二行表頭應為:科別");
            }
            if ("" + cells[0, 2].Value != "姓名")
            {
                errorList.Add("頁籤:畢業生榜單，第三行表頭應為:姓名");
            }
            if ("" + cells[0, 3].Value != "性別")
            {
                errorList.Add("頁籤:畢業生榜單，第四行表頭應為:性別");
            }
            if ("" + cells[0, 4].Value != "班級")
            {
                errorList.Add("頁籤:畢業生榜單，第五行表頭應為:班級");
            }
            if ("" + cells[0, 5].Value != "座號")
            {
                errorList.Add("頁籤:畢業生榜單，第六行表頭應為:座號");
            }
            if ("" + cells[0, 6].Value != "學校")
            {
                errorList.Add("頁籤:畢業生榜單，第七行表頭應為:學校");
            }
            if ("" + cells[0, 7].Value != "學系")
            {
                errorList.Add("頁籤:畢業生榜單，第八行表頭應為:學系");
            } 
            #endregion

            #region 驗證原住民名單
            List<string> stuNumber_Name_list = new List<string>();

            // 整理學號+ _ + 姓名 清單， 最為下面 原住民檢查使用
            foreach (var row in cells.Rows)
            {
                if (row.Index > 0 && !row.IsBlank)
                {
                    //Key: 學號+ _ + 姓名
                    stuNumber_Name_list.Add("" + cells[row.Index, 0].Value + "_" + cells[row.Index, 2].Value);
                }
            }

            if (wb.Worksheets["原住民名單"] == null)
            {
                MsgBox.Show("Excel檔案 沒有'原住民名單' 頁籤，請確認樣板格式正確", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Worksheet ws_abo_list = wb.Worksheets["原住民名單"];

            Cells cells_abo_list = ws_abo_list.Cells;

            if ("" + cells_abo_list[0, 0].Value != "學號")
            {
                errorList.Add("頁籤:原住民名單，第一行表頭應為:學號");
            }
            if ("" + cells_abo_list[0, 1].Value != "姓名")
            {
                errorList.Add("頁籤:原住民名單，第二行表頭應為:姓名");
            }
            if ("" + cells_abo_list[0, 2].Value != "族別")
            {
                errorList.Add("頁籤:原住民名單，第二行表頭應為:族別");
            }

            // 驗證使用者輸入 原住民名單 是否正確
            foreach (var row in cells_abo_list.Rows)
            {
                if (row.Index > 0 && !row.IsBlank)
                {
                    //Key: 學號+ _ + 姓名
                    if (!stuNumber_Name_list.Contains("" + cells_abo_list[row.Index, 0].Value + "_" + cells_abo_list[row.Index, 1].Value))
                    {
                        errorList.Add("頁籤:原住民名單，第" + (row.Index + 1) + "列，學生:" + cells_abo_list[row.Index, 1].Value + " ，其學號+姓名不存在於畢業生榜單上，請確認輸入格式內容正確。");
                    }
                }
            } 
            #endregion


            #endregion

            if (errorList.Count > 0)
            {
                StringBuilder errorMessages = new StringBuilder(); ;

                foreach (string errormessage in errorList)
                {
                    errorMessages.AppendLine(errormessage);
                }

                MsgBox.Show(errorMessages.ToString(), "錯誤訊息", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);

                return;
            }

            // 暫停UI 功能
            textBoxX1.Enabled = false;
            buttonX1.Enabled = false;
            buttonX2.Enabled = false;
            buttonX3.Enabled = false;

            //加入背景執行序
            _bw.DoWork += new DoWorkEventHandler(_bkWork_DoWork);
            _bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_worker_RunWorkerCompleted);
            _bw.ProgressChanged += new ProgressChangedEventHandler(_worker_ProgressChanged);
            _bw.WorkerReportsProgress = true;

            _bw.RunWorkerAsync();

        }

        private void _bkWork_DoWork(object sender, DoWorkEventArgs e)
        {
            // 0.載入匯入檔案
            Workbook wb = new Workbook(source_data);

            Worksheet ws = wb.Worksheets["畢業生榜單"];

            Cells cells = ws.Cells;

            List<string> school_list = new List<string>();

            //整理學校清單
            foreach (var row in cells.Rows)
            {
                if (row.Index > 0 && !row.IsBlank)
                {
                    if (!school_list.Contains("" + cells[row.Index, 6].Value))
                    {
                        school_list.Add("" + cells[row.Index, 6].Value);
                    }
                }
            }

            Workbook wb_school_list = new Workbook(new MemoryStream(Properties.Resources.畢業生進路調查公務統計報表_學校分類_));

            Worksheet ws_school_list = wb_school_list.Worksheets["學校清單"];

            Cells cells_school_list = ws_school_list.Cells;

            // 從第二行開始加
            int row_index = 1;

            int success_count = 0;

            //string formula = "";

            foreach (string school_name in school_list)
            {
                cells_school_list[row_index, 0].Value = school_name;

                //cells_school_list[row_index, 1].Formula = formula;

                cells_school_list[row_index, 0].SetStyle(cells_school_list[1, 0].GetStyle());

                cells_school_list[row_index, 1].SetStyle(cells_school_list[1, 1].GetStyle());

                row_index++;

                try
                {
                    success_count++;

                    int progress = (success_count * 100 / school_list.Count);

                    _bw.ReportProgress(progress);
                }
                catch (Exception ex)
                {
                    MsgBox.Show(ex.Message, "錯誤", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                }
            }


            wb.Worksheets.Add("學校清單");


            // 自整理好的樣板複製過去
            wb.Worksheets["學校清單"].Copy(ws_school_list);


            wb.Worksheets["學校清單"].MoveTo(0);

            wb.CalculateFormula();


            e.Result = wb;
        }


        private void _worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("處理中...", e.ProgressPercentage);
        }

        private void _worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Workbook wb = (Workbook)e.Result;

            SaveFileDialog sd = new SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = "畢業生進路調查公務統計報表(學校分類).xlsx";
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
            FISCA.Presentation.MotherForm.SetStatusBarMessage("畢業生進路調查公務統計報表(學校分類) 產生完成");
            // 任務結束，關閉
            this.Close();
        }

        //private void labelX4_Click(object sender, EventArgs e)
        //{
        //    //Workbook wb = new Workbook(new MemoryStream(Properties.Resources.新編多元性向測驗文華高中測驗範例樣板_CSV));

        //    Workbook wb = new Workbook();

        //    SaveFileDialog sd = new SaveFileDialog();
        //    sd.Title = "另存新檔";
        //    sd.FileName = "新編多元性向測驗文華高中測驗範例樣板.xlsx";
        //    sd.Filter = "Excel檔案 (*.xlsx)|*.xlsx|所有檔案 (*.*)|*.*";
        //    if (sd.ShowDialog() == DialogResult.OK)
        //    {
        //        try
        //        {
        //            wb.Save(sd.FileName);
        //            System.Diagnostics.Process.Start(sd.FileName);
        //        }
        //        catch
        //        {
        //            MessageBox.Show("指定路徑無法存取。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        }
        //    }
        //}
    }
}
