using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Aspose.Words;
using Campus.Report;
using System.IO;
using Counsel_System2.DAO;
using K12.Data;
using System.Xml.Linq;
using Aspose.Words.Reporting;
using FISCA.Presentation.Controls;
using K12.Data.Configuration;
using System.Xml;

namespace Counsel_System2.Forms
{
    public partial class StudInterviewDataReportForm : FISCA.Presentation.Controls.BaseForm
    {
        private string _ReportName = "學生輔導晤談紀錄表_樣版";
        private ReportConfiguration _Config;
        UDTTransfer _UDTTransfer;
        BackgroundWorker _bgWorker;
        List<Dictionary<string, string>> _dataDictList;

        public enum SelectType { 學生, 教師 }

        Dictionary<string, StudentRecord> _studDataDict;
        List<string> _studIDList;
        SelectType _UserSelectType;
        List<string> _TeacherIDList;

        private bool 是否使用範本 { get; set; }

        /// <summary>
        /// 建構子
        /// </summary>
        /// <param name="type"></param>
        public StudInterviewDataReportForm(SelectType type)
        {
            InitializeComponent();
            _UDTTransfer = new UDTTransfer();
            _studIDList = new List<string>();
            _TeacherIDList = new List<string>();
            _studDataDict = new Dictionary<string, StudentRecord>();
            _bgWorker = new BackgroundWorker();
            _dataDictList = new List<Dictionary<string, string>>();
            _UserSelectType = type;
            _bgWorker.DoWork += new DoWorkEventHandler(_bgWorker_DoWork);
            _bgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgWorker_RunWorkerCompleted);
        }

        /// <summary>
        /// FormLoad
        /// </summary>
        private void StudInterviewDataReportForm_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
            _Config = new ReportConfiguration(_ReportName);

            #region 取得上次列印設定值

            ConfigData cd = K12.Data.School.Configuration[_ReportName];
            string Mode = cd["Setup"];
            if (!string.IsNullOrEmpty(Mode))
            {
                bool check = false;
                if (bool.TryParse(Mode, out check))
                {
                    if (check) //如果是bool
                    {
                        rbDEF_2.Checked = true;
                    }
                    else
                    {
                        rbDEF_1.Checked = true;
                    }
                }
                else
                {
                    rbDEF_1.Checked = true;
                }
            }
            else
            {
                rbDEF_1.Checked = true;
            }

            #endregion
        }

        /// <summary>
        /// 開始列印
        /// </summary>
        private void btnPrint_Click(object sender, EventArgs e)
        {
            btnPrint.Enabled = false;

            if (_UserSelectType == SelectType.學生)
                _studIDList = K12.Presentation.NLDPanels.Student.SelectedSource;
            else
                _TeacherIDList = K12.Presentation.NLDPanels.Teacher.SelectedSource;

            是否使用範本 = rbDEF_2.Checked;

            _bgWorker.RunWorkerAsync();

        }

        void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _studDataDict.Clear();
            _dataDictList.Clear();
            string SchoolName = School.ChineseName;

            SaveConfig(是否使用範本);

            Dictionary<int, string> intTeacherNameDict = new Dictionary<int, string>();
            foreach (TeacherRecord tr in Teacher.SelectAll())
            {
                if (tr.Status == TeacherRecord.TeacherStatus.刪除)
                    continue;

                int tid = int.Parse(tr.ID);
                string TName = tr.Name;
                if (!string.IsNullOrWhiteSpace(tr.Nickname))
                    TName = tr.Name + "(" + tr.Nickname + ")";

                intTeacherNameDict.Add(tid, TName);
            }

            if (_UserSelectType == SelectType.學生)
            {
                #region 學生
                // 取得學生資料
                foreach (StudentRecord stud in Student.SelectByIDs(_studIDList))
                    _studDataDict.Add(stud.ID, stud);

                // 取得晤談紀錄
                List<UDT_CounselStudentInterviewRecordDef> dataList = _UDTTransfer.GetCounselStudentInterviewRecordByStudentIDList(_studIDList);

                // 透過 query 依照班級序號、班級名稱、學號座號排序
                _studIDList = Utility.SortStudentID1(_studIDList);

                foreach (string studID in _studIDList)
                {
                    int sid = int.Parse(studID);
                    string studName = "";
                    string StudNumber = "";
                    string ClassName = "";
                    string StudSeatNo = "";
                    if (_studDataDict.ContainsKey(studID))
                    {
                        studName = _studDataDict[studID].Name;
                        StudNumber = _studDataDict[studID].StudentNumber;
                        if (_studDataDict[studID].Class != null)
                            ClassName = _studDataDict[studID].Class.Name;

                        if (_studDataDict[studID].SeatNo.HasValue)
                            StudSeatNo = _studDataDict[studID].SeatNo.Value.ToString();
                    }

                    List<UDT_CounselStudentInterviewRecordDef> studInterviewList = (from da in dataList where da.StudentID == sid orderby da.InterviewDate select da).ToList();

                    foreach (UDT_CounselStudentInterviewRecordDef data in studInterviewList)
                    {
                        Dictionary<string, string> mapDict = new Dictionary<string, string>();
                        mapDict.Add("校名", SchoolName);
                        mapDict.Add("學生姓名", studName);
                        mapDict.Add("學號", StudNumber);
                        mapDict.Add("班級", ClassName);
                        if (intTeacherNameDict.ContainsKey(data.TeacherID))
                            mapDict.Add("晤談老師", intTeacherNameDict[data.TeacherID]);

                        mapDict.Add("晤談編號", data.InterviewNo);
                        mapDict.Add("座號", StudSeatNo);
                        mapDict.Add("晤談對象", data.IntervieweeType);
                        mapDict.Add("晤談方式", data.InterviewType);
                        if (data.InterviewDate.HasValue)
                            mapDict.Add("晤談日期", data.InterviewDate.Value.ToShortDateString());
                        mapDict.Add("時間", data.InterviewTime);
                        mapDict.Add("地點", data.Place);
                        mapDict.Add("參與人員", ParseUDTXML1(data.Attendees));
                        mapDict.Add("晤談事由", data.Cause);
                        mapDict.Add("輔導方式", ParseUDTXML1(data.CounselType));
                        mapDict.Add("輔導歸類", ParseUDTXML1(data.CounselTypeKind));
                        mapDict.Add("內容要點", data.ContentDigest);
                        mapDict.Add("記錄者姓名", data.AuthorName);
                        _dataDictList.Add(mapDict);
                    }

                }
                #endregion
            }

            if (_UserSelectType == SelectType.教師)
            {
                #region 教師
                List<int> sortStudIDList = new List<int>();
                // 透過教師取得所屬晤談紀錄
                List<UDT_CounselStudentInterviewRecordDef> InterviewRecorT = _UDTTransfer.GetCounselStudentInterviewRecordByTeacherIDList(_TeacherIDList);
                // 取得晤談紀錄內學生id
                List<string> InterviewRecorStudIDList = new List<string>();
                foreach (UDT_CounselStudentInterviewRecordDef data in InterviewRecorT)
                {
                    if (!sortStudIDList.Contains(data.StudentID))
                        sortStudIDList.Add(data.StudentID);

                    string key = data.StudentID.ToString();
                    if (!InterviewRecorStudIDList.Contains(key))
                        InterviewRecorStudIDList.Add(key);
                }
                // 取得學生資訊
                Dictionary<int, StudentRecord> InterviewRecorStudDict = new Dictionary<int, StudentRecord>();
                foreach (StudentRecord rec in Student.SelectByIDs(InterviewRecorStudIDList))
                {
                    int sid = int.Parse(rec.ID);
                    if (!InterviewRecorStudDict.ContainsKey(sid))
                        InterviewRecorStudDict.Add(sid, rec);
                }

                // 轉換教師ID int 
                List<int> intTeacherIDList = new List<int>();
                foreach (string tid in _TeacherIDList)
                    intTeacherIDList.Add(int.Parse(tid));

                sortStudIDList = Utility.SortStudentID2(sortStudIDList);
                // 組資料
                foreach (int id in intTeacherIDList)
                {

                    List<UDT_CounselStudentInterviewRecordDef> dataTT = (from data in InterviewRecorT where data.TeacherID == id orderby data.StudentID, data.InterviewDate select data).ToList();
                    // 依班級順序、班級名稱、學生座號排序後加入
                    List<UDT_CounselStudentInterviewRecordDef> dataT = new List<UDT_CounselStudentInterviewRecordDef>();
                    foreach (int ssid in sortStudIDList)
                    {
                        foreach (UDT_CounselStudentInterviewRecordDef data in dataTT.Where(x => x.StudentID == ssid))
                            dataT.Add(data);
                    }

                    foreach (UDT_CounselStudentInterviewRecordDef data in dataT)
                    {
                        string studName = "";
                        string StudNumber = "";
                        string ClassName = "";
                        string StudSeatNo = "";
                        if (InterviewRecorStudDict.ContainsKey(data.StudentID))
                        {
                            studName = InterviewRecorStudDict[data.StudentID].Name;
                            StudNumber = InterviewRecorStudDict[data.StudentID].StudentNumber;
                            if (InterviewRecorStudDict[data.StudentID].Class != null)
                                ClassName = InterviewRecorStudDict[data.StudentID].Class.Name;

                            if (InterviewRecorStudDict[data.StudentID].SeatNo.HasValue)
                                StudSeatNo = InterviewRecorStudDict[data.StudentID].SeatNo.Value.ToString();
                        }

                        Dictionary<string, string> mapDict = new Dictionary<string, string>();
                        mapDict.Add("校名", SchoolName);
                        mapDict.Add("學生姓名", studName);
                        mapDict.Add("學號", StudNumber);
                        mapDict.Add("班級", ClassName);
                        if (intTeacherNameDict.ContainsKey(data.TeacherID))
                            mapDict.Add("晤談老師", intTeacherNameDict[data.TeacherID]);

                        mapDict.Add("晤談編號", data.InterviewNo);
                        mapDict.Add("座號", StudSeatNo);
                        mapDict.Add("晤談對象", data.IntervieweeType);
                        mapDict.Add("晤談方式", data.InterviewType);
                        if (data.InterviewDate.HasValue)
                            mapDict.Add("晤談日期", data.InterviewDate.Value.ToShortDateString());
                        mapDict.Add("時間", data.InterviewTime);
                        mapDict.Add("地點", data.Place);
                        mapDict.Add("參與人員", ParseUDTXML1(data.Attendees));
                        mapDict.Add("晤談事由", data.Cause);
                        mapDict.Add("輔導方式", ParseUDTXML1(data.CounselType));
                        mapDict.Add("輔導歸類", ParseUDTXML1(data.CounselTypeKind));
                        mapDict.Add("內容要點", data.ContentDigest);
                        mapDict.Add("記錄者姓名", data.AuthorName);
                        _dataDictList.Add(mapDict);
                    }
                }
                #endregion
            }

            // word 資料合併
            Document doc = new Document();
            doc.Sections.Clear();

            // 比對欄位名稱放值
            List<string> mapFieldName = new List<string>();
            mapFieldName.Add("校名");
            mapFieldName.Add("學生姓名");
            mapFieldName.Add("學號");
            mapFieldName.Add("班級");
            mapFieldName.Add("晤談老師");
            mapFieldName.Add("晤談編號");
            mapFieldName.Add("座號");
            mapFieldName.Add("晤談對象");
            mapFieldName.Add("晤談方式");
            mapFieldName.Add("晤談日期");
            mapFieldName.Add("時間");
            mapFieldName.Add("地點");
            mapFieldName.Add("參與人員");
            mapFieldName.Add("晤談事由");
            mapFieldName.Add("輔導方式");
            mapFieldName.Add("輔導歸類");
            mapFieldName.Add("內容要點");
            mapFieldName.Add("記錄者姓名");

            //取得範本樣式

            MemoryStream _template;
            if (!是否使用範本)
            {
                _template = new MemoryStream(Properties.Resources.學生晤談記錄表_新範本);
            }
            else
            {
                ConfigData cd = K12.Data.School.Configuration[_ReportName];
                XmlElement config = cd.GetXml("XmlData", null);
                if (config != null)
                {
                    string templateBase64 = config.InnerText;
                    byte[] _buffer = Convert.FromBase64String(templateBase64);
                    _template = new MemoryStream(_buffer);
                }
                else
                {
                    _template = new MemoryStream(Properties.Resources.學生晤談記錄表_新範本);
                }
            }


            foreach (Dictionary<string, string> data in _dataDictList)
            {
                DataTable dt = new DataTable();
                // 建立欄位名稱
                foreach (string name in mapFieldName)
                    dt.Columns.Add(name, typeof(string));

                DataRow dr = dt.NewRow();
                foreach (string name in mapFieldName)
                {
                    if (data.ContainsKey(name))
                        dr[name] = data[name];
                }
                dt.Rows.Add(dr);

                Document docTemplate = new Document(_template);

                docTemplate.MailMerge.FieldMergingCallback = new InsertDocumentAtMailMergeHandler();
                //已過時,改用Aspose.Words.Reporting.MailMergeCleanupOptions.RemoveEmptyParagraphs
                //docTemplate.MailMerge.RemoveEmptyParagraphs = true;

                docTemplate.MailMerge.CleanupOptions = Aspose.Words.Reporting.MailMergeCleanupOptions.RemoveEmptyParagraphs;
                docTemplate.MailMerge.Execute(dt);
                docTemplate.MailMerge.DeleteFields();
                doc.Sections.Add(doc.ImportNode(docTemplate.Sections[0], true));
            }
            e.Result = doc;

        }

        /// <summary>
        /// 記錄設定值
        /// </summary>
        private void SaveConfig(bool check)
        {
            ConfigData cd = K12.Data.School.Configuration[_ReportName];
            cd["Setup"] = check.ToString();
            cd.Save();

            //XmlElement config = cd.GetXml("XmlData", null);
            //if (config == null)
            //{
            //    config = new XmlDocument().CreateElement("XmlData");
            //}
            //string base64 = Convert.ToBase64String(_template.ToArray());
            //config.InnerText = base64;
            //cd.SetXml("XmlData", config);
            //cd.Save();
        }

        void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnPrint.Enabled = true;
            Document doc = (Document)e.Result;
            string _FilePathAndName = "";
            // 當沒有設定檔案名稱
            if (string.IsNullOrEmpty(_FilePathAndName))
            {
                SaveFileDialog sd = new SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = "晤談紀錄表.docx";
                sd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        doc.Save(sd.FileName, Aspose.Words.SaveFormat.Docx);
                        System.Diagnostics.Process.Start(sd.FileName);
                    }
                    catch
                    {
                        MessageBox.Show("指定路徑無法存取。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            else
            {
                try
                {
                    doc.Save(_FilePathAndName, Aspose.Words.SaveFormat.Docx);
                    System.Diagnostics.Process.Start(_FilePathAndName);
                }
                catch
                {
                    SaveFileDialog sd = new SaveFileDialog();
                    sd.Title = "另存新檔";
                    sd.FileName = "晤談紀錄表1.docx";
                    sd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
                    if (sd.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            doc.Save(sd.FileName, Aspose.Words.SaveFormat.Docx);
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
        }

        private class InsertDocumentAtMailMergeHandler : IFieldMergingCallback
        {
            public void FieldMerging(FieldMergingArgs e)
            {
                DocumentBuilder _builder = new DocumentBuilder(e.Document);

            }

            public void ImageFieldMerging(ImageFieldMergingArgs args)
            {

            }
        }

        /// <summary>
        /// 解析 XML 內資料
        /// </summary>
        private string ParseUDTXML1(string strData)
        {
            string str = "";
            List<string> data = new List<string>();
            StringBuilder sb = new StringBuilder();
            sb.Append("<root>");
            sb.Append(strData);
            sb.Append("</root>");
            XElement elmRoot = XElement.Parse(sb.ToString());
            if (elmRoot != null)
            {
                foreach (XElement elm in elmRoot.Elements("Item"))
                {
                    if (elm.Attribute("remark") != null)
                    {
                        string rmk = elm.Attribute("name").Value + ":" + elm.Attribute("remark").Value;
                        data.Add(rmk);
                    }
                    else
                        data.Add(elm.Attribute("name").Value);
                }
            }
            if (data.Count > 0)
                str = string.Join("、", data.ToArray());

            return str;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //檢視預設範本
        private void lnkDownload_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "另存新檔";
            sfd.FileName = "學生晤談記錄表_範本.docx";
            sfd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {

                    FileStream fs = new FileStream(sfd.FileName, FileMode.Create);
                    fs.Write(Properties.Resources.學生晤談記錄表_新範本, 0, Properties.Resources.學生晤談記錄表_新範本.Length);
                    fs.Close();
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch
                {
                    MsgBox.Show("指定路徑無法存取。", "另存檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        //檢視自定範本
        private void linkViewGeDin_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ConfigData cd = K12.Data.School.Configuration[_ReportName];
            XmlElement config = cd.GetXml("XmlData", null);

            if (config != null)
            {
                string templateBase64 = config.InnerText;
                byte[] _buffer = Convert.FromBase64String(templateBase64);

                if (_buffer != null)
                {
                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Title = "另存新檔";
                    sfd.FileName = "自訂學生晤談記錄表範本.docx";
                    sfd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            Aspose.Words.Document doc = new Aspose.Words.Document(new MemoryStream(_buffer));
                            doc.Save(sfd.FileName, Aspose.Words.SaveFormat.Docx);
                        }
                        catch (Exception ex)
                        {
                            MsgBox.Show("檔案無法儲存。" + ex.Message);
                            return;
                        }

                        try
                        {
                            System.Diagnostics.Process.Start(sfd.FileName);
                        }
                        catch (Exception ex)
                        {
                            MsgBox.Show("檔案無法開啟。" + ex.Message);
                            return;
                        }
                    }
                }
                else
                {
                    MsgBox.Show("無自訂範本");
                }
            }
            else
            {
                MsgBox.Show("無自訂範本\n請使用上傳功能上傳範本");
            }
        }

        //上傳新範本
        private void linkUpData_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "選擇自訂的學生晤談記錄表範本";
            ofd.Filter = "Word檔案 (*.docx)|*.docx";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileStream fs = new FileStream(ofd.FileName, FileMode.Open);

                    byte[] tempBuffer = new byte[fs.Length];
                    fs.Read(tempBuffer, 0, tempBuffer.Length);
                    string base64 = Convert.ToBase64String(tempBuffer);
                    fs.Close();

                    ConfigData cd = K12.Data.School.Configuration[_ReportName];
                    XmlElement XmlData = cd.GetXml("XmlData", null);

                    if (XmlData != null)
                    {
                        XmlData.InnerText = base64;
                    }
                    else
                    {
                        XmlDocument xmldoc = new XmlDocument();
                        xmldoc.LoadXml("<XmlData/>");
                        XmlData = (XmlElement)xmldoc.SelectSingleNode("XmlData");
                        XmlData.InnerText = base64;
                    }
                    cd.SetXml("XmlData", XmlData);
                    cd.Save();

                    rbDEF_2.Checked = true;

                    MsgBox.Show("上傳成功。");
                }
                catch
                {
                    MsgBox.Show("指定路徑無法存取。", "開啟檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }
    }
}
