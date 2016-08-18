using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Aspose.Cells;
using Campus.Report;
using Counsel_System2.DAO;
using System.IO;

namespace Counsel_System2.Forms
{
    /// <summary>
    /// 晤談紀錄簽認表
    /// </summary>
    public partial class StudInterviewDataReport1Form : FISCA.Presentation.Controls.BaseForm
    {
        BackgroundWorker _bgWorker;
        private string _ReportName = "高雄輔導晤談紀錄簽認表";
        private ReportConfiguration _Config;
        DateTime _BeginDate, _EndDate;
        string _txtMessage = "";

        List<string> _ErrorList = new List<string>();
        /// <summary>
        /// 教師晤談紀錄
        /// </summary>
        Dictionary<string, List<Rpt_InterviewRecord>> _InterviewRecordDict = new Dictionary<string, List<Rpt_InterviewRecord>>();

        public StudInterviewDataReport1Form()
        {
            InitializeComponent();
            _bgWorker = new BackgroundWorker();
            _bgWorker.DoWork += new DoWorkEventHandler(_bgWorker_DoWork);
            _bgWorker.WorkerReportsProgress = true;
            _bgWorker.ProgressChanged += new ProgressChangedEventHandler(_bgWorker_ProgressChanged);
            _bgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgWorker_RunWorkerCompleted);
        }
        
        void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnPrint.Enabled = true;
            if (e.Error != null)
            {
                FISCA.Presentation.Controls.MsgBox.Show("產生過程發生錯誤:" + e.Error.Message);
                return;
            }

            if (_ErrorList.Count > 0)
            {
                FISCA.Presentation.Controls.MsgBox.Show("內容要點紅色字表示內容字數超過Excel單一儲存格字數512字");
            }

            Workbook wb = (Workbook)e.Result;
            FISCA.Presentation.MotherForm.SetStatusBarMessage("輔導晤談紀錄簽認表產生完成。", 100);

            try
            {
                string FilePath = Application.StartupPath + "\\Reports\\輔導晤談紀錄簽認表.xls";
                wb.Save(FilePath,Aspose.Cells.FileFormatType.Excel2003);
                System.Diagnostics.Process.Start(FilePath);

            }
            catch
            {
                SaveFileDialog sd = new SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = "輔導晤談紀錄簽認表.xls";
                sd.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        wb.Save(sd.FileName, Aspose.Cells.FileFormatType.Excel2003);
                        System.Diagnostics.Process.Start(sd.FileName);
                    }
                    catch
                    {
                        MessageBox.Show("指定路徑無法存取。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
        }

        void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("輔導晤談紀錄簽認表產生中..", e.ProgressPercentage);
        }

        void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorker.ReportProgress(1);
            _ErrorList.Clear();

            // 讀取樣板
            Workbook wbTemplate = new Workbook();
            Workbook wb = new Workbook();
            wbTemplate.Open(new MemoryStream(Properties.Resources.晤談紀錄簽認表樣版));
            wb.Open(new MemoryStream(Properties.Resources.晤談紀錄簽認表樣版));
         
            Range R_Title = wbTemplate.Worksheets[1].Cells.CreateRange(0, 1, false);
            Range R_DescMessage = wbTemplate.Worksheets[1].Cells.CreateRange(1, 1, false);
            Range R_Field = wbTemplate.Worksheets[1].Cells.CreateRange(2, 1, false);
            Range R_Value = wbTemplate.Worksheets[1].Cells.CreateRange(3, 1, false);
            Range R_Last=wbTemplate.Worksheets[1].Cells.CreateRange(5,2,false);

            _bgWorker.ReportProgress(20);

            // 取得晤談紀錄
            _InterviewRecordDict = Utility.GetInterviewRecordDictByTeacherID1(K12.Presentation.NLDPanels.Teacher.SelectedSource);

            int RowTitleIdx = 0, RowDescMsgIdx = 1, RowFieldMsgIdx = 2, RowValueIdx = 3, RowLastIdx = 5, RowCount = 0;

            string TitleName = K12.Data.School.ChineseName + " " + K12.Data.School.DefaultSchoolYear + "學年度第" + K12.Data.School.DefaultSemester + "學期導師晤談紀錄簽認表   " + DateTime.Now.ToShortDateString();
                
            foreach (KeyValuePair<string, List<Rpt_InterviewRecord>> dataList in _InterviewRecordDict)
            {
                RowCount = 0;
                List<Rpt_InterviewRecord> HasDataList =(from da in dataList.Value where da.interview_date >= _BeginDate && da.interview_date <= _EndDate select da).ToList();
                if (HasDataList.Count == 0)
                    continue;

                // 複製樣板表格
                // 校名
                wb.Worksheets[0].Cells.CreateRange(RowTitleIdx, 1, false).Copy(R_Title);
                wb.Worksheets[0].Cells[RowTitleIdx, 0].PutValue(TitleName);

                // 說明
                wb.Worksheets[0].Cells.CreateRange(RowDescMsgIdx, 2, false).Copy(R_DescMessage);
                wb.Worksheets[0].Cells[RowDescMsgIdx, 0].PutValue(_txtMessage);

                wb.Worksheets[0].Cells.CreateRange(RowFieldMsgIdx, 1, false).Copy(R_Field);

                RowValueIdx = RowFieldMsgIdx + 1;
                foreach (Rpt_InterviewRecord data in HasDataList)
                {
                    wb.Worksheets[0].Cells.CreateRange(RowValueIdx, 1, false).Copy(R_Value);
                    // 班級
                    wb.Worksheets[0].Cells[RowValueIdx, 0].PutValue(data.ClassName);
                    // 座號
                    if(data.SeatNo.HasValue)
                        wb.Worksheets[0].Cells[RowValueIdx, 1].PutValue(data.SeatNo.Value);

                    // 姓名
                    wb.Worksheets[0].Cells[RowValueIdx, 2].PutValue(data.Name);
                    // 晤談老師
                    wb.Worksheets[0].Cells[RowValueIdx, 3].PutValue(data.TeacherName);
                    // 晤談方式
                    wb.Worksheets[0].Cells[RowValueIdx,4].PutValue(data.interview_type);
                    // 晤談對象
                    wb.Worksheets[0].Cells[RowValueIdx, 5].PutValue(data.interviewee_type);
                    // 晤談日期
                    wb.Worksheets[0].Cells[RowValueIdx, 6].PutValue(data.interview_date.ToShortDateString());

                    // 內容要點(超過時)
                    if (data.content_digest.Length > 512)
                    {
                        wb.Worksheets[0].Cells[RowValueIdx, 7].Style.Font.Color = Color.Red;
                        _ErrorList.Add("內容要點字數超過");
                    }
                    wb.Worksheets[0].Cells[RowValueIdx, 7].PutValue(data.content_digest);
                                    
                    RowValueIdx++;
                    RowCount++;
                }

                RowLastIdx = RowValueIdx+1;

                wb.Worksheets[0].Cells.CreateRange(RowLastIdx, 2, false).Copy(R_Last);

                RowTitleIdx += RowCount + 7;
                RowDescMsgIdx = RowTitleIdx + 1;
                RowFieldMsgIdx = RowTitleIdx + 2;
                
                wb.Worksheets[0].HPageBreaks.Add(RowLastIdx + 2, 0);
            }
            _bgWorker.ReportProgress(90);            
            wb.Worksheets.RemoveAt(1);
            e.Result = wb;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            btnPrint.Enabled = false;
            
            // 取得使用者畫面選項並儲存
                       

            _BeginDate = dtBegin.Value;
            _EndDate = dtEnd.Value;
            
            if (dtBegin.Value > dtEnd.Value)
            {
                _BeginDate = dtEnd.Value;
                _EndDate = dtBegin.Value;
            }

            _txtMessage = txtDesc.Text;

            _Config.SetString("開始日期", _BeginDate.ToShortDateString());
            _Config.SetString("結束日期", _EndDate.ToShortDateString());
            _Config.SetString("說明文字", _txtMessage);
            _Config.Save();
            _bgWorker.RunWorkerAsync();

        }

        private void StudInterviewDataReport1Form_Load(object sender, EventArgs e)
        {
            // 畫面不大不動
            this.MaximumSize =this.MinimumSize=this.Size;

            // 讀取設定預設值
            _Config = new ReportConfiguration(_ReportName);
            //設定預設值
            SetDefaultTemplate();
        }

        /// <summary>
        /// 設定預設值
        /// </summary>
        private void SetDefaultTemplate()
        {
            DateTime dtb, dte;
            dtEnd.Value=dtBegin.Value = DateTime.Now;
            
            if(DateTime.TryParse(_Config.GetString("開始日期", ""),out dtb))
                dtBegin.Value=dtb;

            if(DateTime.TryParse(_Config.GetString("結束日期", ""),out dte))
                dtEnd.Value=dte;

            txtDesc.Text = _Config.GetString("說明文字", "");
        }
    }
}
