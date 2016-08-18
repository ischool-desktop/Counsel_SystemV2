using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Aspose.Cells;
using System.Xml.Linq;
using K12.Data;

namespace Counsel_System2.Forms
{
    public partial class ABCardPrintForm : FISCA.Presentation.Controls.BaseForm
    {
        BackgroundWorker _bgWork;
        DAO.ABCardPrintManager _ABCardPrintManager;
        DAO.StudentInfoTransfer _StudentInfoTransfer;
        DAO.DocTemplateTransfer _DocTemplateTransfer;
        DAO.DocumentMerge _DocumentMerge;
        List<Dictionary<string, string>> _DataDict;
        bool _fromPrint = false;
        bool _fromTempExport = false;

        List<string> _MappingField;
        string _ReportName = "輔導綜合資料紀錄表樣版";
        public ABCardPrintForm()
        {
            InitializeComponent();

            this.MaximumSize = this.MinimumSize = this.Size;
            _MappingField = new List<string>();
            _bgWork = new BackgroundWorker();
            _bgWork.DoWork += new DoWorkEventHandler(_bgWork_DoWork);
            _bgWork.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgWork_RunWorkerCompleted);
            _ABCardPrintManager = new DAO.ABCardPrintManager();
            _StudentInfoTransfer = new DAO.StudentInfoTransfer();
            _DataDict = new List<Dictionary<string, string>>();
            _DocTemplateTransfer = new DAO.DocTemplateTransfer(_ReportName);
            
        }

        void _bgWork_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

           
            lblMsg.Visible = false;

            if (_fromPrint)
            {
                if (chkSplitByStudentNumber.Checked)
                {
                    FolderBrowserDialog fbd = new FolderBrowserDialog();
                    fbd.Description = "請選擇檔案儲存的資料夾..";

                    if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        foreach (Dictionary<string, string> data in _DataDict)
                        {
                            foreach (KeyValuePair<string, string> val in data)
                            {
                                // 當有學號才處理
                                if (val.Key == "B_學號" && val.Value != "")
                                {
                                    string num = val.Value;
                                    List<Dictionary<string, string>> tmpList = new List<Dictionary<string, string>>();
                                    tmpList.Add(data);
                                    _DocumentMerge = new DAO.DocumentMerge(_MappingField, tmpList, _DocTemplateTransfer.GetUsingTemplate(), fbd.SelectedPath + "\\" + num, false);
                                    _DocumentMerge.Merge();
                                    break;
                                }
                            }
                        }
                        // 是否置自動開啟資料夾
                        if (FISCA.Presentation.Controls.MsgBox.Show("儲存完成,請問是否自動開啟儲存資料夾?", "儲存", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.Yes)
                        {
                            System.Diagnostics.Process proc = new System.Diagnostics.Process();
                            proc.StartInfo.FileName = fbd.SelectedPath;
                            proc.Start();
                        }
                    }
                }
                else
                {
                    // 全部在一個檔案
                      SaveFileDialog sd = new SaveFileDialog();
                    sd.Title = "另存新檔";
                    sd.FileName = "Doc1.doc";
                    sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                    if (sd.ShowDialog() == DialogResult.OK)
                    {
                        bool isOopenFile = true;

                        if (FISCA.Presentation.Controls.MsgBox.Show("請問儲存後是否自動開啟檔案?", "儲存", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.Yes)
                            isOopenFile = true;
                        else
                            isOopenFile = false;

                        _DocumentMerge = new DAO.DocumentMerge(_MappingField, _DataDict, _DocTemplateTransfer.GetUsingTemplate(), "", isOopenFile);
                        if (chkSplitByStudentNumber.Checked == false)
                            _DocumentMerge.Merge();
                    }
                }                

                _fromPrint = false;
                btnPrint.Enabled = true;
            }

            if (_fromTempExport)
            {
                ExportMappingFieldToExcel();
                _fromTempExport = false;
                ilblExportExcelField.Enabled = true;
            }
        }

        void _bgWork_DoWork(object sender, DoWorkEventArgs e)
        {
            _DataDict.Clear();
            _MappingField.Clear();

            List<string> StudentIDList=K12.Presentation.NLDPanels.Student.SelectedSource;

            List<DAO.StudentInfo> StudInfoList=_StudentInfoTransfer.GetStudentInfoList(StudentIDList);

            // 取得畫面所選學生學習歷程
            Global._StudentSemesterHistoryItemDict.Clear();
            foreach (SemesterHistoryRecord shr in SemesterHistory.SelectByStudentIDs(StudentIDList))
            {
                // 依照年級、學期排序後加入                
                Global._StudentSemesterHistoryItemDict.Add(shr.RefStudentID, (from data in shr.SemesterHistoryItems orderby data.GradeYear, data.Semester select data).ToList());
            }

            // 取得日常生活表現具體建議
            Global._ABCard_StudentTextScoreDict = Utility.GetABCard_StudentTextScoreDict(StudentIDList);

            // 畢業資料
            Global._ABCard_StudentGraduateDict = Utility.GetStudentGraduateDict(StudentIDList);

            // 異動資料
            Global._ABCard_StudentUpdateRecDict = Utility.GetStudentUpdateRecDict(StudentIDList);

            //獎懲資料
            Global._ABCard_StudentMDRecordDict = Utility.GetStudentMDRecDict(StudentIDList);

            // 參加社團
            Global._ABCard_StudentSpecCourseDict = Utility.GetABCard_StudentSpecCourseDict(StudentIDList);

            // 擔任幹部
            Global._ABCard_StudentTheCadreDict = Utility.GetABCard_StudentTheCadreDict(StudentIDList);

            // 學生領域成績與畢業成績
            Global._AB_StudSemsDomainScoreDict = Utility.GetAB_StudSemsDomainScoreDict(StudentIDList);

            // 輔導個案會議
            Global._AB_CaseMeetingRecordToABRptDataDict = Utility.ParseCaseMeetingRecordToABRptData(StudentIDList);

            // 輔導優先關懷
            Global._AB_CareRecordToABRptDataDict = Utility.ParseCareRecordToABRptData(StudentIDList);

            // 輔導晤談紀錄
            Global._AB_InterViewDataToABRptDataDict = Utility.ParseInterViewDataToABRptData(StudentIDList);

            // 學生入學照片
            Global._AB_StudentFreshmanDict = Photo.SelectFreshmanPhoto(StudentIDList);

            // 學生畢業照片
            Global._AB_StudentGraduateDict = Photo.SelectGraduatePhoto(StudentIDList);

            // 取得答案
            Dictionary<string, List<DAO.AnswerPkey>> _AnsData = _ABCardPrintManager.GetAnswers(StudentIDList);
            List<string> QuetionFieldList=_ABCardPrintManager.GetMappingField();            
            _MappingField.AddRange(QuetionFieldList);

            Global._StudentNumberList.Clear();
            // studID
            foreach (string studID in StudentIDList)
            {
               
                    Dictionary<string,string> _ansDict = new Dictionary<string,string> ();
                    
                    // 處理學生基本資料、缺曠獎懲、..

                    DAO.StudentInfo studInfo = null;
                    foreach (DAO.StudentInfo si in StudInfoList.Where(x => x.StudentID == studID))
                        studInfo = si;

                    List<string> temp1 = new List<string>();
                    if (studInfo != null)
                    {
                        
                        Dictionary<string, string> baseDict = new Dictionary<string, string>();
                        baseDict.Add("學校中文名稱", studInfo.SchoolName);

                        temp1.Clear();
                        foreach (KeyValuePair<string, string> data in studInfo.AttendanceDict)
                        {
                            baseDict.Add("B_缺曠_" + data.Key, data.Value);
                            temp1.Add(data.Key + "(" + data.Value+")");
                        }

                        if (temp1.Count > 0)
                            baseDict.Add("B_缺曠", string.Join(",", temp1.ToArray()));

                        baseDict.Add("B_入學前學校名稱", studInfo.BeforeEnrollSchoolName);
                        baseDict.Add("B_入學年", studInfo.BeforeEnrollSchoolYear);
                        baseDict.Add("B_出生日期", studInfo.Birthday);
                        baseDict.Add("B_出生日期民國格式", studInfo.BirthdayTW);
                        baseDict.Add("B_出生地", studInfo.Birthplace);
                        baseDict.Add("B_畢業年月", studInfo.GraduationYearMonth);
                        baseDict.Add("B_身分證號", studInfo.IDNumber);
                        baseDict.Add("B_聯絡地址", studInfo.MailingAddress);
                        baseDict.Add("B_聯絡電話", studInfo.MailingPhone);
                        baseDict.Add("B_性別", studInfo.Gender);
                        
                        temp1.Clear();
                        foreach (KeyValuePair<string, string> data in studInfo.MeritDict)
                        {
                            baseDict.Add("B_獎懲_" + data.Key, data.Value);
                            temp1.Add(data.Key + "(" + data.Value+")");
                        }

                        if(temp1.Count>0)
                            baseDict.Add("B_獎懲",string.Join(",",temp1.ToArray()));


                        baseDict.Add("B_姓名", studInfo.Name);
                        baseDict.Add("B_戶籍地址", studInfo.PermanentAddress);
                        baseDict.Add("B_戶籍電話", studInfo.PermanentPhone);
                        baseDict.Add("B_學號", studInfo.StudentNumber);
                        Global._StudentNumberList.Add(studInfo.StudentNumber);
                                                
                        baseDict.Add("GL_異動紀錄",string.Join(";",studInfo.UpdateRecordList.ToArray()));

                        foreach (KeyValuePair<string, string> data in baseDict)
                        {
                            if(!_MappingField.Contains(data.Key))
                                _MappingField.Add(data.Key);

                            _ansDict.Add(data.Key, data.Value);
                        }                        
                    }

                    // 而外加入需要填在這合併名稱，處理在DocumentMerge
                    // 加入學期歷程
                    _ansDict.Add("學習歷程", studID);
                    _ansDict.Add("學期對照", studID);
                    _MappingField.Add("學習歷程");
                    _MappingField.Add("學期對照");

                    // 加入導師評語與具體建議
                    _ansDict.Add("導師評語", studID);
                    _ansDict.Add("具體建議", studID);
                    _MappingField.Add("導師評語");
                    _MappingField.Add("具體建議");

                    // 加入畢業資料
                    _ansDict.Add("畢業資料", studID);
                    _MappingField.Add("畢業資料");

                    // 加入異動資料
                    _ansDict.Add("異動資料", studID);
                    _MappingField.Add("異動資料");

                    // 加入獎懲資料
                    _ansDict.Add("獎懲資料", studID);
                    _MappingField.Add("獎懲資料");

                    // 加入參加社團
                    _ansDict.Add("參加社團", studID);
                    _MappingField.Add("參加社團");

                    // 加入擔任幹部
                    _ansDict.Add("擔任幹部", studID);
                    _MappingField.Add("擔任幹部");

                    // 加入學生領域與畢業成績
                    _ansDict.Add("學習領域與畢業成績", studID);
                    _MappingField.Add("學習領域與畢業成績");

                    // 加入學生輔導個案會議
                    _ansDict.Add("輔導個案會議", studID);
                    _MappingField.Add("輔導個案會議");
                    _MappingField.Add("輔導個案會議其它內容");

                    // 加入學生輔導優先關懷
                    _ansDict.Add("輔導優先關懷", studID);
                    _MappingField.Add("輔導優先關懷");
                    _MappingField.Add("輔導優先關懷其它內容");

                    // 加入學生輔導晤談紀錄
                    _ansDict.Add("輔導晤談紀錄", studID);
                    _MappingField.Add("輔導晤談紀錄");
                    _MappingField.Add("輔導晤談紀錄其它內容");
                
                    // 加入學生入學照片
                    _ansDict.Add("入學照片", studID);
                    _MappingField.Add("入學照片");

                    // 加入學生畢業照片
                    _ansDict.Add("畢業照片", studID);
                    _MappingField.Add("畢業照片");


                    if (_AnsData.ContainsKey(studID))
                    {

                        // 加入測驗紀錄
                        _ansDict.Add("測驗紀錄", studID);
                        _MappingField.Add("測驗紀錄");

                        foreach (string qField in QuetionFieldList)
                        {
                            foreach (DAO.AnswerPkey ap in _AnsData[studID])
                            { //string qqField = ap.Question.GetField();

                                if (ap.Question.QType == "grid")
                                {
                                    string Idx1 = "GR_" + ap.Question.GetField();
                                    if (!_ansDict.ContainsKey(Idx1))
                                    {
                                        _ansDict.Add(Idx1, ap.dataElement.ToString());
                                        if (!_MappingField.Contains(Idx1))
                                            _MappingField.Add(Idx1);
                                    }
                                    int rowIdx = 1;

                                    foreach (XElement elm in ap.dataElement.Elements("Item"))
                                    {
                                        foreach (XElement elm2 in elm.Elements("Field"))
                                        {
                                            string key = ap.Question.GetField() + "_" + elm2.Attribute("key").Value + "R";
                                            if (key == qField)
                                            {
                                                if (!_ansDict.ContainsKey(qField))
                                                {
                                                    string newField = qField + rowIdx;

                                                    if (elm2.Attribute("value") == null)
                                                        _ansDict.Add(newField, "");
                                                    else
                                                        _ansDict.Add(newField, elm2.Attribute("value").Value);

                                                    if (!_MappingField.Contains(newField))
                                                        _MappingField.Add(newField);
                                                }
                                            }
                                        }
                                        rowIdx++;
                                    }

                                }
                                else
                                {
                                    //處理單一分割

                                    string strVal = "";
                                    if (ap.dataElement.Attribute("value") != null)
                                        strVal = ap.dataElement.Attribute("value").Value;

                                    if (ap.Question.GetField() == qField)
                                        if (!_ansDict.ContainsKey(qField))
                                        {
                                            _ansDict.Add(qField, strVal);
                                        }

                                    // 處理組合在一起
                                    if (qField.IndexOf("G") == 0)
                                        if (ap.Question.GetFieldG() == qField)
                                        {
                                            if (strVal != "")
                                                strVal = ap.Question.QLabel + ":" + strVal;

                                            if (!_ansDict.ContainsKey(qField))
                                            {
                                                _ansDict.Add(qField, strVal);
                                            }
                                            else
                                            {
                                                if (strVal != "")
                                                    _ansDict[qField] += "," + strVal;
                                            }
                                        }
                                }
                            }
                        }
                    }
                    _DataDict.Add(_ansDict);
                            
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
                lblMsg.Visible = true;
                btnPrint.Enabled = false;
                _fromPrint = true;
                _bgWork.RunWorkerAsync();            
        }

        private void ABCardPrintForm_Load(object sender, EventArgs e)
        {
            lblMsg.Visible = false;
        }

        private void ilblUploadTemplate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            _DocTemplateTransfer.UploadTemplate();
        }

        private void ilblDownTemplate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            _DocTemplateTransfer.DownloadTemplate();
        }

        private void ilbDefaultTemplate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            _DocTemplateTransfer.RemoveAndUseDefaultTemplate();
        }

        private void ilblExportExcelField_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            lblMsg.Visible = true;
            ilblExportExcelField.Enabled = false;
            _fromTempExport = true;
            _bgWork.RunWorkerAsync();

            
        }

        public void ExportMappingFieldToExcel()
        {
            Workbook wb = new Workbook();
            wb.Worksheets[0].Name = "輔導綜合資料合併欄位";
            _MappingField.Sort();
            // 取得匯出資料
            if (_MappingField.Count > 255)
            {
                FISCA.Presentation.Controls.MsgBox.Show("超過Excel能放的欄位255欄，無法使用Word合併欄位功能，系統將資料轉成直視.");

                int RowIdx = 0;
                foreach (string name in _MappingField)
                {
                    wb.Worksheets[0].Cells[RowIdx, 0].PutValue(name);
                    RowIdx++;
                }

            }
            else
            {
                // 放入 Workbook
                int ColIdx = 0;
                foreach (string name in _MappingField)
                {
                    wb.Worksheets[0].Cells[0, ColIdx].PutValue(name);
                    ColIdx++;
                }
            }

            // 存檔
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "Excel (*.xls)|*.xls";
            if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    wb.Save(saveDialog.FileName);
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("儲存失敗" + ex.Message);
                    return;
                }

                try
                {
                    System.Diagnostics.Process.Start(saveDialog.FileName);
                }

                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("開啟失敗" + ex.Message);
                    return;
                }
            }        
        }

        private void groupPanel1_Click(object sender, EventArgs e)
        {

        }
    }
}
