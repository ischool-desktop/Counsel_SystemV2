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
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace Counsel_System2.Forms
{
    public partial class StudABCardReportForm : FISCA.Presentation.Controls.BaseForm
    {
        private string _ReportName = "學生輔導資料紀錄表AB表範本";
        private ReportConfiguration _Config;
        public enum SelectType { 學生, 班級 }
        public enum SelectTemplateType {預設,自訂 }
        public enum SelectFileType {單檔,學號分檔 }
        // 有錯誤記錄
        string _ErrMsg1 = "";
        List<string> _ErrorList = new List<string>();
        // 樣板內合併欄位名稱
        Dictionary<string, byte> _TemplateFieldDict = new Dictionary<string, byte>();


        // 最大資料筆數,data table columns 使用
        int _MaxDataCols = 30;

        // 使用者選擇列印是學生還是班級
        SelectType _UserSelectType;
        // 使用者選的樣板
        SelectTemplateType _SelectTemplateType;
        // 使用者選的檔案
        SelectFileType _SelectFileType;
        UDTTransfer _UDTTransfer;
        BackgroundWorker _bgWorker;
        DataTable _dtTable;

        List<string> _StudentIDList;


        // 學生基本資料暫存
        Dictionary<string, StudentRecord> _StudRecDict = new Dictionary<string, StudentRecord>();

        // 學習歷程
        Dictionary<string, SemesterHistoryRecord> _SemesterHistoryRecordDict = new Dictionary<string, SemesterHistoryRecord>();
        
        // 地址暫存
        Dictionary<string, AddressRecord> _AddressRecordDict = new Dictionary<string, AddressRecord>();

        // 電話暫存
        Dictionary<string, PhoneRecord> _PhoneRecordDict = new Dictionary<string, PhoneRecord>();

        // 異動
        Dictionary<string, List<UpdateRecordRecord>> _UpdateRecordRecordDict = new Dictionary<string, List<UpdateRecordRecord>>();
        // 異動細項
        Dictionary<string, XElement> _UpdateRecordRecordInfoDict = new Dictionary<string, XElement>();
        // AutoSummary
        Dictionary<string, List<K12.BusinessLogic.AutoSummaryRecord>> _AutoSummaryRecordDict = new Dictionary<string, List<K12.BusinessLogic.AutoSummaryRecord>>();

        // 文字評量(國高中位置不同)
        Dictionary<string, List<ABCard_StudentTextScore>> _StudentTextScoreDict = new Dictionary<string, List<ABCard_StudentTextScore>>();
        
        // 缺曠暫存
        Dictionary<string, List<AttendanceRecord>> _AttendanceRecordDict = new Dictionary<string, List<AttendanceRecord>>();

        // 獎懲暫存
        Dictionary<string, List<DisciplineRecord>> _DisciplineRecordDict = new Dictionary<string, List<DisciplineRecord>>();

        // 入學照片
        Dictionary<string, string> _FreshmanPhotoDict = new Dictionary<string, string>();

        // 學期成績
        Dictionary<string, List<SemesterScoreRecord>> _SemesterScoreRecordDict = new Dictionary<string, List<SemesterScoreRecord>>();
        // 國中學期領域成績
        Dictionary<string, XElement> _SemesterDomainScoreDict = new Dictionary<string, XElement>();
        // 高中分項成績
        Dictionary<string, List<AB_SchoolYearSemesterIdx>> _SemesterEnrtyScoreSchoolYearDict = new Dictionary<string, List<AB_SchoolYearSemesterIdx>>();
        Dictionary<string, XElement> _SemesterEnrtyScoreDict = new Dictionary<string, XElement>();

        // 國中畢業成績
        Dictionary<string, Dictionary<string, decimal?>> _GradeScoreDict = new Dictionary<string, Dictionary<string, decimal?>>();

        // 高中畢業成績
        Dictionary<string, XElement> _GradeScoreSHDict = new Dictionary<string, XElement>();

        // 畢業照片
        Dictionary<String, string> _GraduatePhotoDict = new Dictionary<string, string>();

        // 心理測驗題目
        Dictionary<int, UDT_QuizDef> _QuizDefDict = new Dictionary<int, UDT_QuizDef>();

        // 心理測驗答案
        Dictionary<string, List<UDT_StudQuizDataDef>> _StudQuizDataDict = new Dictionary<string, List<UDT_StudQuizDataDef>>();

        // 晤談紀錄
        Dictionary<string, List<UDT_CounselStudentInterviewRecordDef>> _CounselStudentInterviewRecordDict = new Dictionary<string, List<UDT_CounselStudentInterviewRecordDef>>(); 

        // 個案會議
        Dictionary<string, List<UDT_CounselCaseMeetingRecordDef>> _CounselCaseMeetingRecordDict = new Dictionary<string, List<UDT_CounselCaseMeetingRecordDef>>();

        // 優先關懷
        Dictionary<string, List<UDT_CounselCareRecordDef>> _CounselCareRecordDict = new Dictionary<string, List<UDT_CounselCareRecordDef>>();

        // 綜合紀錄表-單值
        Dictionary<string, List<UDTSingleRecordDef>> _SingleRecordDict = new Dictionary<string, List<UDTSingleRecordDef>>();

        // 綜合紀錄表-多值
        Dictionary<string, List<UDTMultipleRecordDef>> _MultipleRecordDict = new Dictionary<string, List<UDTMultipleRecordDef>>();

        // 綜合紀錄表-學期
        Dictionary<string, List<UDTSemesterDataDef>> _SemesterDataDict = new Dictionary<string, List<UDTSemesterDataDef>>();

        // 綜合紀錄表-學年
        Dictionary<string, List<UDTYearlyDataDef>> _YearlyDataDict = new Dictionary<string, List<UDTYearlyDataDef>>();

        // 綜合紀錄表-直系血親
        Dictionary<string, List<UDTRelativeDef>> _RelativeDict = new Dictionary<string, List<UDTRelativeDef>>();
        
        // 綜合紀錄表-兄弟姊妹
        Dictionary<string, List<UDTSiblingDef>> _SiblingDict = new Dictionary<string, List<UDTSiblingDef>>();

        // 綜合紀錄表-優先順序
        Dictionary<string, List<UDTPriorityDataDef>> _PriorityDataDict = new Dictionary<string, List<UDTPriorityDataDef>>();

        public StudABCardReportForm(SelectType selType,List<string> StudentIDList)
        {
            
            InitializeComponent();
            _UserSelectType = selType;
            _UDTTransfer = new UDTTransfer();
            _bgWorker = new BackgroundWorker();
            _dtTable = new DataTable();
            _StudentIDList = StudentIDList;
            _bgWorker.WorkerReportsProgress = true;
            _bgWorker.DoWork += new DoWorkEventHandler(_bgWorker_DoWork);
            _bgWorker.ProgressChanged += new ProgressChangedEventHandler(_bgWorker_ProgressChanged);
            _bgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgWorker_RunWorkerCompleted);
        }

        void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("輔導資料紀錄表產生中", e.ProgressPercentage);
        }

        void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // 產生報表
            btnPrint.Enabled = true;

            if (e.Error != null)
            {
                FISCA.Presentation.Controls.MsgBox.Show("產生過程發生錯誤:" + e.Error.Message);
                return;
            }
           

            FISCA.Presentation.MotherForm.SetStatusBarMessage("輔導資料紀錄表產生完成。", 100);

            if (_ErrorList.Count > 0)
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("資料合併欄位不足，會產生部分資料無法呈現：");
                foreach (string str in _ErrorList)
                    sb.AppendLine(str);

                StudABCardReportForm_ErrorMsg errMsg = new StudABCardReportForm_ErrorMsg(sb);
                errMsg.ShowDialog();
            }            

            Document document = (Document)e.Result;
            string inputReportName = "輔導資料紀錄表";
            string reportName = inputReportName;

            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".doc");

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
                document.Save(path, Aspose.Words.SaveFormat.Doc);
                System.Diagnostics.Process.Start(path);

               
            }
            catch
            {
                System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".doc";
                sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        document.Save(sd.FileName, Aspose.Words.SaveFormat.Doc);
                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return;
                    }
                }
            }
        }

        void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            UDTTransfer _UDTTransfer = new UDTTransfer();
            // 清空暫存 table
            _dtTable.Clear();
            _dtTable.Columns.Clear();
            _ErrorList.Clear();
            // 新增欄位
            AddTableSColumn();

            Document docTemplae = new Document();

            // 檢查所使用樣板
            if(_SelectTemplateType == SelectTemplateType.預設)
                docTemplae = new Document(new MemoryStream(Properties.Resources.輔導資料紀錄表範本));

            if (_SelectTemplateType == SelectTemplateType.自訂)
            {
                //docTemplae = _Config.Template.ToDocument();
                docTemplae = new Document(new MemoryStream(_Config.Template.ToBinary()));                
            }  

            _bgWorker.ReportProgress(1);
            // 加入樣板內合併欄位名稱
            _TemplateFieldDict.Clear();
            foreach (string name in docTemplae.MailMerge.GetFieldNames())
                if (!_TemplateFieldDict.ContainsKey(name))
                    _TemplateFieldDict.Add(name, 1);


            // word 資料合併
            Document doc = new Document();
            doc.Sections.Clear();

            // 讀取需要資料並建立索引
            _StudRecDict.Clear();
            foreach (StudentRecord rec in Student.SelectByIDs(_StudentIDList))
                _StudRecDict.Add(rec.ID, rec);

            // 學習歷程
            _SemesterHistoryRecordDict.Clear();
            foreach (SemesterHistoryRecord rec in SemesterHistory.SelectByStudentIDs(_StudentIDList))
            {
                if (!_SemesterHistoryRecordDict.ContainsKey(rec.RefStudentID))
                    // 排序
                    rec.SemesterHistoryItems = (from data in rec.SemesterHistoryItems orderby data.SchoolYear, data.Semester select data).ToList();
                    _SemesterHistoryRecordDict.Add(rec.RefStudentID, rec);
            }

            // 建立地址
            _AddressRecordDict.Clear();
            foreach (AddressRecord rec in Address.SelectByStudentIDs(_StudentIDList))
                _AddressRecordDict.Add(rec.RefStudentID, rec);

            // 建立電話
            _PhoneRecordDict.Clear();
            foreach (PhoneRecord rec in Phone.SelectByStudentIDs(_StudentIDList))
                _PhoneRecordDict.Add(rec.RefStudentID, rec);

            // 建立異動
            _UpdateRecordRecordDict.Clear();            
            foreach (UpdateRecordRecord rec in UpdateRecord.SelectByStudentIDs(_StudentIDList))
            {
                if (!_UpdateRecordRecordDict.ContainsKey(rec.StudentID))
                    _UpdateRecordRecordDict.Add(rec.StudentID, new List<UpdateRecordRecord>());

                _UpdateRecordRecordDict[rec.StudentID].Add(rec);

            }
            _UpdateRecordRecordInfoDict.Clear();
            _UpdateRecordRecordInfoDict = Utility.GetUpdateRecordInfo(_StudentIDList);

            // 建立AutoSummary
            _AutoSummaryRecordDict.Clear();
            foreach (K12.BusinessLogic.AutoSummaryRecord rec in K12.BusinessLogic.AutoSummary.Select(_StudentIDList, null))
            {
                if (!_AutoSummaryRecordDict.ContainsKey(rec.RefStudentID))
                    _AutoSummaryRecordDict.Add(rec.RefStudentID, new List<K12.BusinessLogic.AutoSummaryRecord>());

                _AutoSummaryRecordDict[rec.RefStudentID].Add(rec);
            }

            // 建立文字評量
            _StudentTextScoreDict.Clear();
            _StudentTextScoreDict = Utility.GetABCard_StudentTextScoreDict(_StudentIDList);

            //// 建立缺曠
            //_AttendanceRecordDict.Clear();
            //foreach (AttendanceRecord rec in Attendance.SelectByStudentIDs(_StudentIDList))
            //{
            //    if (!_AttendanceRecordDict.ContainsKey(rec.RefStudentID))
            //        _AttendanceRecordDict.Add(rec.RefStudentID, new List<AttendanceRecord>());

            //    _AttendanceRecordDict[rec.RefStudentID].Add(rec);
            //}

            // 學期
            _SemesterScoreRecordDict.Clear();
            foreach (SemesterScoreRecord data in SemesterScore.SelectByStudentIDs(_StudentIDList))
            {                
                if (!_SemesterScoreRecordDict.ContainsKey(data.RefStudentID))
                    _SemesterScoreRecordDict.Add(data.RefStudentID,new List<SemesterScoreRecord> ());

                _SemesterScoreRecordDict[data.RefStudentID].Add(data);
            }
            
            _SemesterDomainScoreDict.Clear();
            // 國中
            _SemesterDomainScoreDict = Utility.GetSemeterDomainScoreByStudentIDList(_StudentIDList);
            // 高中
            _SemesterEnrtyScoreDict = Utility.GetSemeterEntryScoreByStudentIDList(_StudentIDList);
            _SemesterEnrtyScoreSchoolYearDict = Utility.GetSemeterSchoolYearScoreByStudentIDList(_StudentIDList);

            // 國中畢業成績
            _GradeScoreDict.Clear();
            _GradeScoreDict = Utility.GetStudGraduateDictJH(_StudentIDList);

            // 高中畢業
            _GradeScoreSHDict.Clear();
            _GradeScoreSHDict = Utility.GetGradeScoreSHByStudentIDList(_StudentIDList);

            // 建立獎懲明細
            _DisciplineRecordDict.Clear();
            foreach (DisciplineRecord rec in Discipline.SelectByStudentIDs(_StudentIDList))
            {
                if(!_DisciplineRecordDict.ContainsKey(rec.RefStudentID))
                    _DisciplineRecordDict.Add(rec.RefStudentID, new List<DisciplineRecord>());

                _DisciplineRecordDict[rec.RefStudentID].Add(rec);
            }
            // 入學照片
            _FreshmanPhotoDict.Clear();
            _FreshmanPhotoDict = K12.Data.Photo.SelectFreshmanPhoto(_StudentIDList);

            // 畢業照片
            _GraduatePhotoDict.Clear();
            _GraduatePhotoDict = K12.Data.Photo.SelectGraduatePhoto(_StudentIDList);
                        
            // 學校名稱
            string SchoolName = K12.Data.School.ChineseName;

            // 輔導資料
            // 建立心理測驗
            _QuizDefDict.Clear();
            foreach(UDT_QuizDef da in _UDTTransfer.GetAllQuizData())
            {
                int id=int.Parse(da.UID);
                _QuizDefDict.Add(id, da);
            }

            _StudQuizDataDict.Clear();
            foreach (UDT_StudQuizDataDef data in _UDTTransfer.GetStudQuizDataByStudentIDList(_StudentIDList))
            {
                string sid = data.StudentID.ToString();
                if (!_StudQuizDataDict.ContainsKey(sid))
                    _StudQuizDataDict.Add(sid, new List<UDT_StudQuizDataDef>());

                _StudQuizDataDict[sid].Add(data);
            }

            // 建立晤談記錄
            _CounselStudentInterviewRecordDict.Clear();
            foreach (UDT_CounselStudentInterviewRecordDef data in _UDTTransfer.GetCounselStudentInterviewRecordByStudentIDList(_StudentIDList))
            {
                string key = data.StudentID.ToString();
                if (!_CounselStudentInterviewRecordDict.ContainsKey(key))
                    _CounselStudentInterviewRecordDict.Add(key, new List<UDT_CounselStudentInterviewRecordDef>());

                _CounselStudentInterviewRecordDict[key].Add(data);            
            }

            // 個案會議
            _CounselCaseMeetingRecordDict.Clear();
            foreach (UDT_CounselCaseMeetingRecordDef data in _UDTTransfer.GetCaseMeetingRecordListByStudentIDList(_StudentIDList))
            {
                string key = data.StudentID.ToString();
                if (!_CounselCaseMeetingRecordDict.ContainsKey(key))
                    _CounselCaseMeetingRecordDict.Add(key, new List<UDT_CounselCaseMeetingRecordDef>());

                _CounselCaseMeetingRecordDict[key].Add(data);
            }

            // 優先關懷
            _CounselCareRecordDict.Clear();
            foreach (UDT_CounselCareRecordDef data in _UDTTransfer.GetCareRecordsByStudentIDList(_StudentIDList))
            {
                string key = data.StudentID.ToString();
                if (!_CounselCareRecordDict.ContainsKey(key))
                    _CounselCareRecordDict.Add(key, new List<UDT_CounselCareRecordDef>());

                _CounselCareRecordDict[key].Add(data);
            }

            // 綜合紀錄表-單值
            _SingleRecordDict.Clear();
            foreach (UDTSingleRecordDef data in UDTTransfer.ABUDTSingleRecordSelectByStudentIDList(_StudentIDList))
            {
                string key = data.StudentID.ToString();
                if (!_SingleRecordDict.ContainsKey(key))
                    _SingleRecordDict.Add(key, new List<UDTSingleRecordDef>());

                _SingleRecordDict[key].Add(data);

            }

            // 綜合紀錄表-多值
            _MultipleRecordDict.Clear();
            foreach (UDTMultipleRecordDef data in UDTTransfer.ABUDTMultipleRecordSelectByStudentIDList(_StudentIDList))
            {
                string key = data.StudentID.ToString();
                if (!_MultipleRecordDict.ContainsKey(key))
                    _MultipleRecordDict.Add(key, new List<UDTMultipleRecordDef>());

                _MultipleRecordDict[key].Add(data);
            }
            
            // 綜合紀錄表-學期
            _SemesterDataDict.Clear();
            foreach (UDTSemesterDataDef data in UDTTransfer.ABUDTSemesterDataSelectByStudentIDList(_StudentIDList))
            {
                string key = data.StudentID.ToString();
                if (!_SemesterDataDict.ContainsKey(key))
                    _SemesterDataDict.Add(key, new List<UDTSemesterDataDef>());

                _SemesterDataDict[key].Add(data);
            }

            // 綜合紀錄表-學年
            _YearlyDataDict.Clear();
            foreach (UDTYearlyDataDef data in UDTTransfer.ABUDTYearlyDataSelectByStudentIDList(_StudentIDList))
            {
                string key = data.StudentID.ToString();
                if (!_YearlyDataDict.ContainsKey(key))
                    _YearlyDataDict.Add(key, new List<UDTYearlyDataDef>());

                _YearlyDataDict[key].Add(data);
            }

            // 綜合紀錄表-直系血親
            _RelativeDict.Clear();
            foreach (UDTRelativeDef data in UDTTransfer.ABUDTRelativeSelectByStudentIDList(_StudentIDList))
            {
                string key = data.StudentID.ToString();
                if (!_RelativeDict.ContainsKey(key))
                    _RelativeDict.Add(key, new List<UDTRelativeDef>());

                _RelativeDict[key].Add(data);
            }

            // 綜合紀錄表-兄弟姊妹
            _SiblingDict.Clear();
            foreach (UDTSiblingDef data in UDTTransfer.ABUDTSiblingSelectByStudentIDList(_StudentIDList))
            {
                string key = data.StudentID.ToString();
                if (!_SiblingDict.ContainsKey(key))
                    _SiblingDict.Add(key, new List<UDTSiblingDef>());

                _SiblingDict[key].Add(data);
            }

            // 綜合紀錄表-優先順序
            _PriorityDataDict.Clear();
            foreach (UDTPriorityDataDef data in UDTTransfer.ABUDTPriorityDataSelectByStudentIDList(_StudentIDList))
            {
                string key = data.StudentID.ToString();
                if (!_PriorityDataDict.ContainsKey(key))
                    _PriorityDataDict.Add(key, new List<UDTPriorityDataDef>());

                _PriorityDataDict[key].Add(data);
            }

            _bgWorker.ReportProgress(30);

            // 儲存獎懲明細使用
            StringBuilder sbDisciplineRecord = new StringBuilder();

            // 開始填入資料
            int idx = 1;
            foreach (string StudID in _StudentIDList)
            {
                _ErrMsg1 = "";
                DataRow row = _dtTable.NewRow();
                                
                row["學校名稱"] = SchoolName;
                                
                row["姓名"] = "";
                row["性別"] = "";
                row["學號"] = "";
                row["身分證號"] = "";
                row["生日"] = "";
                row["生日2"] = "";
                row["出生地"] = "";
                if (_StudRecDict.ContainsKey(StudID))
                {
                    _ErrMsg1=_StudRecDict[StudID].Name+"：";
                    row["姓名"] = _StudRecDict[StudID].Name;
                    row["性別"] = _StudRecDict[StudID].Gender;
                    row["學號"] = _StudRecDict[StudID].StudentNumber;
                    row["身分證號"] = _StudRecDict[StudID].IDNumber;
                    if (_StudRecDict[StudID].Birthday.HasValue)
                    {                       
                        row["生日"] = _StudRecDict[StudID].Birthday.Value.ToShortDateString();
                        // 100/1/1
                        row["生日2"] = (_StudRecDict[StudID].Birthday.Value.Year - 1911) + "/" + _StudRecDict[StudID].Birthday.Value.Month + "/" + _StudRecDict[StudID].Birthday.Value.Day;
                    }
                    row["出生地"] = _StudRecDict[StudID].BirthPlace;
                }

                row["入學照片"] = "";
                row["入學照片2"] = "";

                if (_FreshmanPhotoDict.ContainsKey(StudID))
                {
                    row["入學照片"] = _FreshmanPhotoDict[StudID];
                    row["入學照片2"] = _FreshmanPhotoDict[StudID];
                }
                if(_GraduatePhotoDict.ContainsKey(StudID))
                {
                    row["畢業照片"] = _GraduatePhotoDict[StudID];
                    row["畢業照片2"] = _GraduatePhotoDict[StudID];
                }
                
                row["戶籍地址"] = "";
                row["聯絡地址"] = "";
                row["戶籍電話"] = "";
                row["聯絡電話"] = "";
                if (_AddressRecordDict.ContainsKey(StudID))
                {
                    row["戶籍地址"] = _AddressRecordDict[StudID].PermanentAddress;
                    row["聯絡地址"] = _AddressRecordDict[StudID].MailingAddress;
                }
                if (_PhoneRecordDict.ContainsKey(StudID))
                {
                    row["戶籍電話"] = _PhoneRecordDict[StudID].Permanent;
                    row["聯絡電話"] = _PhoneRecordDict[StudID].Contact;
                }


                idx = 1;
                // 學習歷程
                foreach (SemesterHistoryItem item in _SemesterHistoryRecordDict[StudID].SemesterHistoryItems)
                {
                    // 檢查料是否超過可合併欄位
                    ChechMapFieldName("學習歷程_學年度", idx);
                    ChechMapFieldName("學習歷程_學期", idx);
                    ChechMapFieldName("學習歷程_班級", idx);
                    ChechMapFieldName("學習歷程_班導師", idx);

                    row["學習歷程_學年度" + idx] = item.SchoolYear;
                    row["學習歷程_學期" + idx] = item.Semester;
                    row["學習歷程_班級" + idx] = item.ClassName;
                    row["學習歷程_班導師" + idx] = item.Teacher;
                    idx++;
                }

                row["入學年月"] = "";
                row["入學學校"] = "";
                row["畢業年月"] = "";
                row["學籍核准文號"] = "";
                row["畢業證書字號"] = "";
                idx = 1;
                // 異動高中
                if (_UpdateRecordRecordDict.ContainsKey(StudID) && _UpdateRecordRecordDict[StudID].Count>0)
                {                   
                    
                    DateTime dtLast = DateTime.MinValue;
                    int lastUID = 0;
                    foreach (UpdateRecordRecord rec in _UpdateRecordRecordDict[StudID])
                    {
                        if (rec.UpdateCode.Length < 3)
                            continue;
                        
                        int code = int.Parse(rec.UpdateCode);
                        DateTime dt = DateTime.Parse(rec.UpdateDate);
                        
                        // 新生
                        if (code < 100)
                        {
                            row["入學年月"] = dt.Year + string.Format("{0:00}", dt.Month);

                            if (_UpdateRecordRecordInfoDict[rec.ID].Element("GraduateSchool") != null)
                                row["入學學校"] = _UpdateRecordRecordInfoDict[rec.ID].Element("GraduateSchool").Value;  // 畢

                            if(row["學籍核准文號"].ToString()=="")
                                row["學籍核准文號"] = rec.ADNumber;

                            //row["入學學校"] = rec.;  // 畢業國中
                        }
                        else if (code == 501)
                        {
                            // 畢業
                            row["畢業年月"] = dt.Year + string.Format("{0:00}", dt.Month);
                            //row["畢業證書字號"] = rec.GraduateCertificateNumber;
                        }
                        else
                        {
                            // 找最後一筆
                            DateTime dtL;
                            if (DateTime.TryParse(rec.LastADDate, out dtL))
                            {
                                if (dtLast == null)
                                    dtLast = dtL;
                                int uid = int.Parse(rec.ID);
                                if (dtL >=dtLast)
                                {
                                    if (uid > lastUID)
                                    {
                                        row["學籍核准文號"] = rec.LastADNumber;
                                        dtLast = dtL;
                                        lastUID = uid;
                                    }
                                }
                            }
                            // 檢查料是否超過可合併欄位
                            ChechMapFieldName("異動日期", idx);
                            ChechMapFieldName("異動原因", idx);
                            row["異動日期" + idx] = dt.ToShortDateString();
                            row["異動原因" + idx] = rec.UpdateDescription;
                            idx++;
                        }
                    }
                }

                // 異動國中                
                if (_UpdateRecordRecordDict.ContainsKey(StudID) && _UpdateRecordRecordDict[StudID].Count > 0)
                {
                    DateTime dtLast1 = DateTime.MinValue;
                    int lastUID1=0;
                    foreach (UpdateRecordRecord rec in _UpdateRecordRecordDict[StudID])
                    {
                        if (rec.UpdateCode.Length > 1)
                            continue;

                        int code = int.Parse(rec.UpdateCode);
                        DateTime dt = DateTime.Parse(rec.UpdateDate);

                        // 新生
                        if (code == 1)
                        {
                            if (_UpdateRecordRecordInfoDict.ContainsKey(rec.ID))
                            {
                                if (_UpdateRecordRecordInfoDict[rec.ID].Element("EnrollmentSchoolYear") != null)
                                    row["入學年月"] = _UpdateRecordRecordInfoDict[rec.ID].Element("EnrollmentSchoolYear").Value;
                                if (_UpdateRecordRecordInfoDict[rec.ID].Element("GraduateSchool") != null)
                                    row["入學學校"] = _UpdateRecordRecordInfoDict[rec.ID].Element("GraduateSchool").Value;  // 畢

                                if (row["學籍核准文號"].ToString() == "")
                                    row["學籍核准文號"] = rec.ADNumber;
                            }
                        }
                        else if (code == 2)
                        {
                            // 畢業
                            if (_UpdateRecordRecordInfoDict.ContainsKey(rec.ID))
                            {
                                if (_UpdateRecordRecordInfoDict[rec.ID].Element("GraduateSchoolYear") != null)
                                    row["畢業年月"] = _UpdateRecordRecordInfoDict[rec.ID].Element("GraduateSchoolYear").Value;

                                if (_UpdateRecordRecordInfoDict[rec.ID].Element("GraduateCertificateNumber") != null)
                                    row["畢業證書字號"] = _UpdateRecordRecordInfoDict[rec.ID].Element("GraduateCertificateNumber").Value;
                            }
                            //row["畢業年月"] = rec.GraduateSchoolYear;
                            //row["畢業證書字號"] = rec.GraduateCertificateNumber;
                        }
                        else
                        {
                            //// 找最後一筆
                            //DateTime dtL;
                            //if (DateTime.TryParse(rec.LastADDate, out dtL))
                            //{
                            //    if (dtLast1 == null)
                            //        dtLast1 = dtL;

                            //    if (dtL >= dtLast1)
                            //    {
                            //        row["學籍核准文號"] = rec.LastADNumber;
                            //        dtLast1 = dtL;
                            //    }
                            //}
                            // 檢查料是否超過可合併欄位
                            ChechMapFieldName("異動日期", idx);
                            ChechMapFieldName("異動原因", idx);
                            row["異動日期" + idx] = dt.ToShortDateString();
                            row["異動原因" + idx] = rec.UpdateDescription;
                            idx++;
                        }
                        // 找最後一筆
                        DateTime dtL;
                        if (DateTime.TryParse(rec.LastADDate, out dtL))
                        {
                            int uid=int.Parse(rec.ID);
                            if (dtLast1 == null)
                                dtLast1 = dtL;

                            if (dtL >= dtLast1)
                            {
                                if (uid > lastUID1)
                                {
                                    row["學籍核准文號"] = rec.LastADNumber;
                                    dtLast1 = dtL;
                                    lastUID1 = uid;
                                }
                            }
                        }
                    }
                }


                if (_SemesterScoreRecordDict.ContainsKey(StudID))
                {
                    List<SemesterScoreRecord> dataList = (from da in _SemesterScoreRecordDict[StudID] orderby da.SchoolYear, da.Semester select da).ToList();
                    int i = 1;
                    foreach (SemesterScoreRecord rec in dataList)
                    {
                        // 國中成績
                        if (_SemesterDomainScoreDict.ContainsKey(rec.ID))
                        {

                            // 檢查料是否超過可合併欄位
                            ChechMapFieldName("學年學期", i);
                            ChechMapFieldName("學期成績", i);
                            List<string> strList = new List<string>();
                            
                            row["學年學期" + i] = rec.SchoolYear + "學年第" + rec.Semester + "學期";


                            if (_SemesterDomainScoreDict[rec.ID].Element("Domains") != null)
                            {
                                foreach (XElement elm in _SemesterDomainScoreDict[rec.ID].Element("Domains").Elements("Domain"))
                                {
                                    string str = elm.Attribute("領域").Value + "：" + elm.Attribute("成績").Value;
                                    strList.Add(str);
                                }
                            }
                            if (strList.Count > 0)
                                row["學期成績" + i] = string.Join("，", strList.ToArray());
                            i++;
                        }
                    }
                }
                    // 高中分項
                    if (_SemesterEnrtyScoreSchoolYearDict.ContainsKey(StudID))
                    {
                        List<AB_SchoolYearSemesterIdx> dataList = (from da in _SemesterEnrtyScoreSchoolYearDict[StudID] orderby da.SchoolYear, da.Semester select da).ToList();
                        int i = 1;
                        foreach (AB_SchoolYearSemesterIdx rec in dataList)
                        {
                            if (_SemesterEnrtyScoreDict.ContainsKey(rec.id))
                            {                                
                                    if (_SemesterEnrtyScoreDict[rec.id].Elements("Entry").Count() > 0)
                                    {
                                        // 檢查料是否超過可合併欄位
                                        ChechMapFieldName("學年學期", i);
                                        ChechMapFieldName("學期成績", i);
                                        List<string> strList = new List<string>();

                                        row["學年學期" + i] = rec.SchoolYear + "學年第" + rec.Semester + "學期";

                                        foreach (XElement elm in _SemesterEnrtyScoreDict[rec.id].Elements("Entry"))
                                        {
                                            string str = elm.Attribute("分項").Value + "：" + elm.Attribute("成績").Value;
                                            strList.Add(str);
                                        }

                                        if (strList.Count > 0)
                                            row["學期成績" + i] = string.Join("，", strList.ToArray());
                                        i++;
                                    }
                            }
                        }
                    }
                
               

                // 國中畢業成績
                if (_GradeScoreDict.ContainsKey(StudID))
                {
                    if (_GradeScoreDict[StudID].Count > 0)
                    {
                        List<string> strList = new List<string>();

                        foreach (KeyValuePair<string, decimal?> data in _GradeScoreDict[StudID])
                        {
                            if (data.Value.HasValue)
                                strList.Add(data.Key + "：" + data.Value.Value);
                        }
                        row["畢業成績"] = string.Join("，", strList.ToArray());
                    }
                }

                // 高中畢業成績
                if (_GradeScoreSHDict.ContainsKey(StudID))
                {
                    if (_GradeScoreSHDict[StudID].Elements("EntryScore").Count() > 0)
                    {
                        List<string> strList = new List<string>();
                        foreach (XElement elm in _GradeScoreSHDict[StudID].Elements("EntryScore"))
                        {
                            if (elm.Attribute("Entry").Value == "德行")
                                continue;

                            string str = elm.Attribute("Entry").Value + "：" + elm.Attribute("Score").Value;
                            strList.Add(str);
                        }
                        row["畢業成績"] = string.Join("，", strList.ToArray());
                    }
                }

                row["導師評語"] = "";
                // 文字評量
                if (_StudentTextScoreDict.ContainsKey(StudID))
                {
                    List<string> strList = new List<string>();
                    foreach (ABCard_StudentTextScore sts in _StudentTextScoreDict[StudID])
                    {
                        string str = "";
                        switch (sts.GradeYear)
                        {
                            case 1: str = "一"; break;
                            case 2: str = "二"; break;
                            case 3: str = "三"; break;
                            case 4: str = "四"; break;
                            case 5: str = "五"; break;
                            case 6: str = "六"; break;
                            case 7: str = "一"; break;
                            case 8: str = "二"; break;
                            case 9: str = "三"; break;
                        }

                        if (str != "")
                        {
                            if (sts.Semester == 1)
                                str += "上：";

                            if (sts.Semester == 2)
                                str += "下：";
                            strList.Add(str + sts.sb_Comment + sts.DailyLifeRecommend);
                        }
                        
                    }
                    if (strList.Count > 0)
                        row["導師評語"] = string.Join("\n", strList.ToArray());
                }
                
                // 處理AutoSummary
                if (_AutoSummaryRecordDict.ContainsKey(StudID))
                {
                    List<K12.BusinessLogic.AutoSummaryRecord> dataList = (from da in _AutoSummaryRecordDict[StudID] orderby da.SchoolYear, da.Semester select da).ToList();
                    int i=1;
                    foreach (K12.BusinessLogic.AutoSummaryRecord data in dataList)
                    {
                        List<string> strList = new List<string> ();
                        foreach (K12.BusinessLogic.AbsenceCountRecord rec in data.AbsenceCounts)
                            strList.Add(rec.PeriodType+"："+rec.Name + rec.Count + "節");

                        if (strList.Count > 0)
                        {
                            // 檢查料是否超過可合併欄位
                            ChechMapFieldName("缺曠統計", i);
                            row["缺曠統計" + i] =data.SchoolYear+"學年第"+data.Semester+"學期："+ string.Join("，", strList.ToArray());
                            i++;
                        }
                    }
                }

                // 獎懲
                if (_DisciplineRecordDict.ContainsKey(StudID))
                {
                    sbDisciplineRecord.Clear();

                    int i=1,ia = 1,ib=1;
                    List<DisciplineRecord> dataList = (from da in _DisciplineRecordDict[StudID] orderby da.OccurDate select da).ToList();
                    foreach (DisciplineRecord data in dataList)
                    {
                        string str = "";
                        // 檢查料是否超過可合併欄位
                        if (!_TemplateFieldDict.ContainsKey("獎懲明細"))
                        {
                            ChechMapFieldName("獎懲日期", i);
                            ChechMapFieldName("獎懲類別支數", i);
                            ChechMapFieldName("獎懲事由", i);
                        }
                        string spstr1 = "";
                        string spstr2 = "";
                        string spstr3 = "";

                        if (data.MeritFlag == "0" || data.MeritFlag == "2")
                        {                           
                            str = "";
                            if (data.DemeritA.HasValue && data.DemeritA.Value>0)
                                str += "大過 " + data.DemeritA.Value;

                            if (data.DemeritB.HasValue && data.DemeritB.Value>0)
                                str += "小過 " + data.DemeritB.Value;

                            if (data.DemeritC.HasValue && data.DemeritC.Value>0)
                                str += "警告 " + data.DemeritC.Value;

                            row["懲戒日期" + ia] = data.OccurDate.ToShortDateString();
                            row["懲戒類別支數" + ia] = str;
                            row["懲戒事由" + ia] = data.Reason;

                            row["獎懲日期" + i] = data.OccurDate.ToShortDateString();
                            row["獎懲類別支數" + i] = str;
                            row["獎懲事由" + i] = data.Reason;

                            spstr1 = data.OccurDate.ToShortDateString();
                            spstr2 = str;
                            spstr3 = data.Reason;

                            if (data.MeritFlag == "2")
                            {
                                row["懲戒類別支數" + ia] = "留校察看";
                                row["獎懲類別支數" + i] = "留校察看";
                                spstr2 = "留校察看";
                            }
                            if (data.ClearDate.HasValue)
                            {
                                row["懲戒日期" + ia] = data.ClearDate.Value.ToShortDateString();
                                row["懲戒類別支數" + ia] = "銷過";
                                row["懲戒事由" + ia] = data.ClearReason;

                                row["獎懲日期" + i] = data.ClearDate.Value.ToShortDateString();
                                row["獎懲類別支數" + i] = "銷過";
                                row["獎懲事由" + i] = data.ClearReason;

                                spstr1 = data.ClearDate.Value.ToShortDateString();
                                spstr2 = "銷過";
                                spstr3 = data.ClearReason;
                            }
                            i++;
                            ia++;
                        }
                        if (data.MeritFlag == "1")
                        {
                            str = "";
                            if (data.MeritA.HasValue && data.MeritA.Value >0)
                                str += "大功 " + data.MeritA.Value;

                            if (data.MeritB.HasValue && data.MeritB.Value>0)
                                str += "小功 " + data.MeritB.Value;

                            if (data.MeritC.HasValue && data.MeritC.Value >0)
                                str += "嘉獎 " + data.MeritC.Value;

                            row["獎勵日期" + ib] = data.OccurDate.ToShortDateString();
                            row["獎勵類別支數" + ib] = str;
                            row["獎勵事由" + ib] = data.Reason;

                            row["獎懲日期" + i] = data.OccurDate.ToShortDateString();
                            row["獎懲類別支數" + i] = str;
                            row["獎懲事由" + i] = data.Reason;

                            spstr1 = data.OccurDate.ToShortDateString();
                            spstr2 = str;
                            spstr3 = data.Reason;

                            i++;
                            ib++;
                        }
                        sbDisciplineRecord.AppendLine(spstr1 + " " + spstr2 + " " + spstr3);
                    }
                    row["獎懲明細"] = sbDisciplineRecord.ToString();
                }

                // 單值
                if (_SingleRecordDict.ContainsKey(StudID))
                {
                    foreach (UDTSingleRecordDef data in _SingleRecordDict[StudID])
                    {
                        switch (data.Key)
                        {
                            case "本人概況_血型": row["血型"] = data.Data; break;
                            case "本人概況_宗教": row["宗教"] = data.Data; break;

                            case "本人概況_原住民血統": 
                                if (data.Data == "有")
                                {
                                    row["原住民血統"] = "■ 有 □ 無";
                                    string[] str = data.Remark.Split('_');
                                    row["原住民血統_稱謂"] = str[0];
                                    row["原住民血統_族別"] = str[1];
                                }
                                else
                                {
                                    row["原住民血統"] = "□ 有 ■ 無";
                                    row["原住民血統_稱謂"] = "    ";
                                    row["原住民血統_族別"] = "    ";
                                }
                                break;

                            case "家庭狀況_監護人_姓名": row["監護人_姓名"] = data.Data; break;
                            case "家庭狀況_監護人_關係": row["監護人_關係"] = data.Data; break;
                            case "家庭狀況_監護人_通訊地址": row["監護人_通訊地址"] = data.Data; break;
                            case "家庭狀況_監護人_電話": row["監護人_電話"] = data.Data; break;

                            case "自傳_家中最了解我的人": row["家中最了解我的人"] = data.Data; break;
                            case "自傳_常指導我做功課的人": row["常指導我做功課的人"] = data.Data; break;
                            case "自傳_讀過且印象最深刻的課外書": row["讀過且印象最深刻的課外書"] = data.Data; break;
                            case "自傳_喜歡的人": row["喜歡的人"] = data.Data; break;
                            case "自傳_喜歡的人_因為": row["喜歡的人_因為"] = data.Data; break;
                            case "自傳_最要好的朋友": row["最要好的朋友"] = data.Data; break;
                            case "自傳_他是怎樣的人": row["他是怎樣的人"] = data.Data; break;
                            case "自傳_最喜歡做的事": row["最喜歡做的事"] = data.Data; break;
                            case "自傳_最喜歡做的事_因為":row["最喜歡做的事_因為"] = data.Data; break;
                            case "自傳_最不喜歡做的事": row["最不喜歡做的事"] = data.Data; break;
                            case "自傳_最不喜歡做的事_因為": row["最不喜歡做的事_因為"] = data.Data; break;
                            case "自傳_國中時的學校生活": row["國中時的學校生活"] = data.Data; break;
                            case "自傳_最快樂的回憶": row["最快樂的回憶"] = data.Data; break;
                            case "自傳_最痛苦的回憶": row["最痛苦的回憶"] = data.Data; break;
                            case "自傳_最足以描述自己的幾句話": row["最足以描述自己的幾句話"] = data.Data; break;
                                //Cloud新增
                            case "自傳_家中最了解我的人_因為": row["家中最了解我的人_因為"] = data.Data; break;
                            case "自傳_我在家中最怕的人是": row["我在家中最怕的人是"] = data.Data; break;
                            case "自傳_我在家中最怕的人是_因為": row["我在家中最怕的人是_因為"] = data.Data; break;
                            case "自傳_我覺得我的優點是": row["我覺得我的優點是"] = data.Data; break;
                            case "自傳_我覺得我的缺點是": row["我覺得我的缺點是"] = data.Data; break;
                            case "自傳_最喜歡的國小（國中）老師": row["最喜歡的國小（國中）老師"] = data.Data; break;
                            case "自傳_最喜歡的國小（國中）老師__因為": row["最喜歡的國小（國中）老師__因為"] = data.Data; break;
                            case "自傳_小學（國中）老師或同學常說我是": row["小學（國中）老師或同學常說我是"] = data.Data; break;
                            case "自傳_小學（國中）時我曾在班上登任過的職務有": row["小學（國中）時我曾在班上登任過的職務有"] = data.Data; break;
                            case "自傳_我在小學（國中）得過的獎有": row["我在小學（國中）得過的獎有"] = data.Data; break;
                            case "自傳_我覺得我自己的過去最滿意的是": row["我覺得我自己的過去最滿意的是"] = data.Data; break;
                            case "自傳_我排遣休閒時間的方法是": row["我排遣休閒時間的方法是"] = data.Data; break;
                            case "自傳_我最難忘的一件事是": row["我最難忘的一件事是"] = data.Data; break;
                            case "自傳_自我的心聲_一年級_我目前遇到最大的困難是": row["自我的心聲_一年級_我目前遇到最大的困難是"] = data.Data; break;
                            case "自傳_自我的心聲_一年級_我目前最需要的協助是": row["自我的心聲_一年級_我目前最需要的協助是"] = data.Data; break;
                            case "自傳_自我的心聲_二年級_我目前遇到最大的困難是": row["自我的心聲_二年級_我目前遇到最大的困難是"] = data.Data; break;
                            case "自傳_自我的心聲_二年級_我目前最需要的協助是": row["自我的心聲_二年級_我目前最需要的協助是"] = data.Data; break;
                            case "自傳_自我的心聲_三年級_我目前遇到最大的困難是": row["自我的心聲_三年級_我目前遇到最大的困難是"] = data.Data; break;
                            case "自傳_自我的心聲_三年級_我目前最需要的協助是": row["自我的心聲_三年級_我目前最需要的協助是"] = data.Data; break;

                            case "自我認識_需要改進的地方_1": row["自我認識_需要改進的地方_1"] = data.Data; break;
                            case "自我認識_優點_1": row["自我認識_優點_1"] = data.Data; break;
                            case "自我認識_個性_1": row["自我認識_個性_1"] = data.Data; break;
                            case "自我認識_需要改進的地方_2": row["自我認識_需要改進的地方_2"] = data.Data; break;
                            case "自我認識_優點_2": row["自我認識_優點_2"] = data.Data; break;
                            case "自我認識_個性_2": row["自我認識_個性_2"] = data.Data; break;
                            case "自我認識_需要改進的地方_3": row["自我認識_需要改進的地方_3"] = data.Data; break;
                            case "自我認識_優點_3": row["自我認識_優點_3"] = data.Data; break;
                            case "自我認識_個性_3": row["自我認識_個性_3"] = data.Data; break;
                            case "生活感想_內容3_1": row["生活感想_內容3_1"] = data.Data; break;
                            case "生活感想_內容2_1": row["生活感想_內容2_1"] = data.Data; break;
                            case "生活感想_內容1_1": row["生活感想_內容1_1"] = data.Data; break;
                            case "生活感想_內容3_2": row["生活感想_內容3_2"] = data.Data; break;
                            case "生活感想_內容2_2": row["生活感想_內容2_2"] = data.Data; break;
                            case "生活感想_內容1_2": row["生活感想_內容1_2"] = data.Data; break;
                            case "備註_備註": row["備註_備註"] = data.Data; break;
                        }
                    }
                }

                // 學年型
                if (_YearlyDataDict.ContainsKey(StudID))
                {
                    // 組合方式一
                    foreach(UDTYearlyDataDef data in _YearlyDataDict[StudID])
                    {
                        switch (data.Key)
                        {
                            case "家庭狀況_父母關係": row["父母關係"] = Utility.Counsel_Yearly_Parse1(data); break;
                            case "家庭狀況_家庭氣氛": row["家庭氣氛"] = Utility.Counsel_Yearly_Parse1(data); break;
                            case "家庭狀況_父親管教方式": row["父管教方式"] = Utility.Counsel_Yearly_Parse1(data); break;
                            case "家庭狀況_母親管教方式": row["母管教方式"] = Utility.Counsel_Yearly_Parse1(data); break;
                            case "家庭狀況_居住環境":row["居住環境"] = Utility.Counsel_Yearly_Parse1(data); break;
                            case "家庭狀況_本人住宿":row["本人住宿"] = Utility.Counsel_Yearly_Parse1(data); break;
                            case "家庭狀況_經濟狀況": row["經濟狀況"] = Utility.Counsel_Yearly_Parse1(data); break;
                            case "家庭狀況_每星期零用錢": row["零用金"] = Utility.Counsel_Yearly_Parse1(data); break;
                            case "家庭狀況_每星期零用錢(元)": row["零用金"] = Utility.Counsel_Yearly_Parse1(data); break;
                            case "家庭狀況_我覺得是否足夠":row["零用金是否足夠"] = Utility.Counsel_Yearly_Parse1(data); break;
                            case "學習狀況_特殊專長": row["特殊專長"] = Utility.Counsel_Yearly_Parse1(data); break;
                            case "學習狀況_休閒興趣": row["休閒興趣"] = Utility.Counsel_Yearly_Parse1(data); break;
                            case "學習狀況_最喜歡的學科": row["最喜歡的學科"] = Utility.Counsel_Yearly_Parse1(data); break;
                            case "學習狀況_最感困難的學科": row["最感困難的學科"] = Utility.Counsel_Yearly_Parse1(data); break;
                        }
                    }
                
                    // 組合方式二                    
                    foreach (UDTYearlyDataDef data in _YearlyDataDict[StudID])
                    {
                        switch (data.Key)
                        {
                            case "適應情形_生活習慣":
                                for( int i=1;i<=6;i++)
                                    row["生活習慣"+i] = Utility.Counsel_Yearly_Parse2(data, i); break;

                            case "適應情形_人際關係":
                                for (int i = 1; i <= 6; i++)
                                    row["人際關係"+i] = Utility.Counsel_Yearly_Parse2(data,i); break;
                            case "適應情形_外向行為":
                                for (int i = 1; i <= 6; i++)
                                row["外向行為"+i] = Utility.Counsel_Yearly_Parse2(data,i); break;
                            case "適應情形_內向行為":
                                for (int i = 1; i <= 6; i++)
                                row["內向行為"+i] = Utility.Counsel_Yearly_Parse2(data, i); break;
                            case "適應情形_學習動機":
                                for (int i = 1; i <= 6; i++)
                                row["學習動機"+i] = Utility.Counsel_Yearly_Parse2(data, i); break;
                            case "適應情形_服務熱忱":
                                for (int i = 1; i <= 6; i++)
                                row["服務熱忱"+i] = Utility.Counsel_Yearly_Parse2(data, i); break;
                            case "適應情形_人生態度":
                                for (int i = 1; i <= 6; i++)
                                row["人生態度"+i] = Utility.Counsel_Yearly_Parse2(data, i); break;
                        }
                    }
                }
                
                // 多值
                if (_MultipleRecordDict.ContainsKey(StudID))
                {
                    row["生理狀態"] = Utility.Counsel_Multiple_Parse(_MultipleRecordDict[StudID], "本人概況_生理缺陷");
                    row["特殊疾病"] = Utility.Counsel_Multiple_Parse(_MultipleRecordDict[StudID], "本人概況_曾患特殊疾病");

                    row["升學意願"] = Utility.Counsel_Multiple_Parse(_MultipleRecordDict[StudID], "畢業後計畫_升學意願");
                    row["就業意願"] = Utility.Counsel_Multiple_Parse(_MultipleRecordDict[StudID], "畢業後計畫_就業意願");
                    row["職訓種類"] = Utility.Counsel_Multiple_Parse(_MultipleRecordDict[StudID], "畢業後計畫_參加職業訓練");
                    row["受訓地區"] = Utility.Counsel_Multiple_Parse(_MultipleRecordDict[StudID], "畢業後計畫_受訓地區");
                }
                // 優先順序
                if (_PriorityDataDict.ContainsKey(StudID))
                {
                    foreach (UDTPriorityDataDef data in _PriorityDataDict[StudID])
                    {
                        switch (data.Key)
                        {
                            case "畢業後計畫_將來職業": row["職業意願"] = Utility.Counsel_PriorityData_Parse1(data); break;
                            case "畢業後計畫_就業地區": row["就業地區"] = Utility.Counsel_PriorityData_Parse1(data); break;
                        }
                    }
                }              

                // 直系血親
                if (_RelativeDict.ContainsKey(StudID))
                {
                    int i = 1;
                    foreach (UDTRelativeDef data in _RelativeDict[StudID])
                    {
                        // 檢查料是否超過可合併欄位
                        ChechMapFieldName("家長親屬_稱謂", i);
                        ChechMapFieldName("家長親屬_姓名", i);
                        ChechMapFieldName("家長親屬_出生年", i);
                        ChechMapFieldName("家長親屬_存歿", i);
                        ChechMapFieldName("家長親屬_教育程度", i);
                        ChechMapFieldName("家長親屬_職業", i);
                        ChechMapFieldName("家長親屬_工作機構", i);
                        ChechMapFieldName("家長親屬_職稱", i);
                        ChechMapFieldName("家長親屬_原國籍", i);

                        row["家長親屬_稱謂" + i] = data.Title;
                        row["家長親屬_姓名" + i] = data.Name;
                        row["家長親屬_出生年" + i] = data.BirthYear;
                        row["家長親屬_存歿" + i]="存";
                        if (data.IsAlive.HasValue && data.IsAlive.Value == false)
                            row["家長親屬_存歿" + i] = "歿";
                        row["家長親屬_教育程度" + i] = data.EduDegree;
                        row["家長親屬_職業" + i] = data.Job;
                        row["家長親屬_工作機構" + i] = data.Institute;
                        row["家長親屬_職稱" + i] = data.JobTitle;
                        row["家長親屬_原國籍" + i] = data.National;
                        i++;
                    }                
                }

                // 兄弟姊妹
                if (_SiblingDict.ContainsKey(StudID))
                {
                    int i = 1;
                    foreach (UDTSiblingDef data in _SiblingDict[StudID])
                    {
                        // 檢查料是否超過可合併欄位
                        ChechMapFieldName("兄弟姊妹_稱謂", i);
                        ChechMapFieldName("兄弟姊妹_姓名", i);
                        ChechMapFieldName("兄弟姊妹_出生年", i);
                        ChechMapFieldName("兄弟姊妹_畢肄業學校", i);

                        row["兄弟姊妹_稱謂" + i] = data.Title;
                        row["兄弟姊妹_姓名" + i] = data.Name;
                        row["兄弟姊妹_出生年" + i] = data.BirthYear;
                        row["兄弟姊妹_畢肄業學校" + i] = data.SchoolName;
                        i++;
                    }                    
                }

                // 學期型
                if (_SemesterDataDict.ContainsKey(StudID))
                {
                    foreach (UDTSemesterDataDef data in _SemesterDataDict[StudID])
                    {
                        switch (data.Key)
                        { 
                            case "本人概況_身高":
                                for (int i = 1; i <= 6; i++)
                                    row["身高" + i] = Utility.Counsel_SemesterData_Parse1(data,i);break;
                            case "本人概況_體重":
                                for (int i = 1; i <= 6; i++)
                                    row["體重" + i] = Utility.Counsel_SemesterData_Parse1(data,i);break;
                        }
                    }
                }


                // 心理測驗
                if (_StudQuizDataDict.ContainsKey(StudID))
                {
                    int i = 1;
                    List<UDT_StudQuizDataDef> dataList = (from da in _StudQuizDataDict[StudID] orderby da.ImplementationDate select da).ToList();
                    foreach (UDT_StudQuizDataDef data in dataList)
                    {
                        // 檢查料是否超過可合併欄位
                        ChechMapFieldName("測驗名稱", i);
                        ChechMapFieldName("測驗日期", i);
                        ChechMapFieldName("測驗結果", i);

                        row["測驗名稱" + i] = "";
                        if (_QuizDefDict.ContainsKey(data.QuizID))                        
                            row["測驗名稱" + i] = _QuizDefDict[data.QuizID].QuizName;
                        if (data.ImplementationDate.HasValue)
                            row["測驗日期" + i] = data.ImplementationDate.Value.ToShortDateString();
                        else
                            row["測驗日期" + i] = "";

                        row["測驗結果" + i] = Utility.CounselStudQuizXmlParse1(data.Content);
                        
                        i++;
                    }                
                }

                // 晤談紀錄
                if (_CounselStudentInterviewRecordDict.ContainsKey(StudID))
                {
                    int i1 = 1;
                    List<UDT_CounselStudentInterviewRecordDef> dataList = (from da in _CounselStudentInterviewRecordDict[StudID] orderby da.InterviewDate select da).ToList();
                    foreach (UDT_CounselStudentInterviewRecordDef data in dataList)
                    {
                        // 檢查料是否超過可合併欄位
                        ChechMapFieldName("晤談紀錄日期", i1);
                        ChechMapFieldName("晤談紀錄對象", i1);
                        ChechMapFieldName("晤談紀錄方式", i1);
                        ChechMapFieldName("晤談紀錄內容要點", i1);
                        ChechMapFieldName("晤談紀錄記錄者姓名", i1);

                        row["晤談紀錄日期"+i1]=data.InterviewDate.Value.ToShortDateString();
                        row["晤談紀錄對象" + i1] = data.IntervieweeType;
                        row["晤談紀錄方式" + i1] = data.InterviewType;
                        row["晤談紀錄內容要點" + i1] = data.ContentDigest;
                        row["晤談紀錄記錄者姓名" + i1] = data.AuthorName;
                        i1++;
                    }
                
                }

                // 個案會議
                if (_CounselCaseMeetingRecordDict.ContainsKey(StudID))
                {
                    int i2 = 1;
                    List<UDT_CounselCaseMeetingRecordDef> dataList = (from da in _CounselCaseMeetingRecordDict[StudID] orderby da.MeetingDate select da).ToList();
                    foreach (UDT_CounselCaseMeetingRecordDef data in dataList)
                    {
                        // 檢查料是否超過可合併欄位
                        ChechMapFieldName("個案會議會議日期", i2);
                        ChechMapFieldName("個案會議會議事由", i2);                        
                        ChechMapFieldName("個案會議內容要點", i2);
                        ChechMapFieldName("個案會議記錄者姓名", i2);

                        row["個案會議會議日期" + i2] = data.MeetingDate.Value.ToShortDateString();
                        row["個案會議會議事由" + i2] = data.MeetingCause;
                        row["個案會議內容要點" + 2] = data.ContentDigest;
                        row["個案會議記錄者姓名" + i2] = data.AuthorName;

                        i2++;
                    }
                }


                // 優先關懷
                if (_CounselCareRecordDict.ContainsKey(StudID))
                {
                    int i3 = 1;
                    List<UDT_CounselCareRecordDef> dataList = (from da in _CounselCareRecordDict[StudID] orderby da.FileDate select da).ToList();
                    foreach (UDT_CounselCareRecordDef data in dataList)
                    {
                        // 檢查料是否超過可合併欄位
                        ChechMapFieldName("優先關懷立案日期", i3);
                        ChechMapFieldName("優先關懷個案類別", i3);
                        ChechMapFieldName("優先關懷個案來源", i3);
                        ChechMapFieldName("優先關懷記錄者姓名", i3);

                        row["優先關懷立案日期" + i3] = data.FileDate.Value.ToShortDateString();
                        row["優先關懷個案類別" + i3] = data.CaseCategory;
                        row["優先關懷個案來源" + i3] = data.CaseOrigin;
                        row["優先關懷記錄者姓名" + i3] = data.AuthorName;
                        i3++;
                    }
                }


                _dtTable.Rows.Add(row);               
            }
            Document document = new Document();
            document = docTemplae;
            doc.MailMerge.FieldMergingCallback = new InsertDocumentAtMailMergeHandler();
            doc.Sections.Add(doc.ImportNode(document.Sections[0], true));
            
            //doc.MailMerge.MergeField += new Aspose.Words.Reporting.MergeFieldEventHandler(MailMerge_MergeField);
            
            doc.MailMerge.Execute(_dtTable);
            doc.MailMerge.RemoveEmptyParagraphs = true;
            doc.MailMerge.DeleteFields();            
            _bgWorker.ReportProgress(95);
            e.Result = doc;
        }

        private class InsertDocumentAtMailMergeHandler : IFieldMergingCallback
        {
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs e)
            {
                if (e.FieldName == "入學照片" || e.FieldName == "入學照片2" || e.FieldName == "畢業照片" || e.FieldName == "畢業照片2")
                {
                    if (e.FieldValue != null && e.FieldValue.ToString() != "")
                    {
                        byte[] photo = Convert.FromBase64String(e.FieldValue.ToString()); //e.FieldValue as byte[];

                        if (photo != null && photo.Length > 0)
                        {
                            DocumentBuilder photoBuilder = new DocumentBuilder(e.Document);
                            photoBuilder.MoveToField(e.Field, true);
                            e.Field.Remove();
                            Shape photoShape = new Shape(e.Document, ShapeType.Image);
                            photoShape.ImageData.SetImage(photo);
                            photoShape.WrapType = WrapType.Inline;
                            Cell cell = photoBuilder.CurrentParagraph.ParentNode as Cell;
                            //cell.CellFormat.LeftPadding = 0;
                            //cell.CellFormat.RightPadding = 0;
                            if (e.FieldName == "入學照片" || e.FieldName == "畢業照片")
                            {
                                // 1吋
                                photoShape.Width = ConvertUtil.MillimeterToPoint(25);
                                photoShape.Height = ConvertUtil.MillimeterToPoint(35);
                            }
                            else
                            {
                                //2吋
                                photoShape.Width = ConvertUtil.MillimeterToPoint(35);
                                photoShape.Height = ConvertUtil.MillimeterToPoint(45);
                            }
                            photoBuilder.InsertNode(photoShape);
                        }
                    }
                }
            }

            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs e)
            {
 
            }
        }

        /// <summary>
        /// 檢查合併資料是否有合併欄位
        /// </summary>
        /// <param name="name"></param>
        /// <param name="i"></param>
        private void ChechMapFieldName(string name, int i)
        {
            string key = name + i;
            if (!_TemplateFieldDict.ContainsKey(key))
                _ErrorList.Add(_ErrMsg1+key);
        }

        //void MailMerge_MergeField(object sender, Aspose.Words.Reporting.MergeFieldEventArgs e)
        //{
           
        //}

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void StudABCardReportForm_Load(object sender, EventArgs e)
        {
            // 控制畫面大小不動
            this.MaximumSize = this.MinimumSize = this.Size;

            // 讀取畫面設定預設值與設定畫面
            _Config = new ReportConfiguration(_ReportName);
            SetDefaultTemplate();
        }

        /// <summary>
        /// 設定預設樣版
        /// </summary>
        private void SetDefaultTemplate()
        {
            if (_Config.Template == null)
            {
                ReportTemplate rptTmp = new ReportTemplate(Properties.Resources.輔導資料紀錄表範本, TemplateType.Word);
                _Config.Template = rptTmp;
                _Config.Save();
            }

            // 讀取設定
            string strTemp = _Config.GetString("SelectTemplateType", "預設");
            string strFile = _Config.GetString("SelectFileType", "單檔");

            if (strTemp == "預設")
                chkDefault.Checked=true;
            else
                chkUserDef.Checked = true;

            if (strFile == "單檔")
                chkFileAllInOne.Checked = true;
            else
                chkFileSplitBySNum.Checked = true;

        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            btnPrint.Enabled = false;
            // 傳入用使用者畫面所選
            if (chkDefault.Checked)
                _SelectTemplateType = SelectTemplateType.預設;

            if (chkUserDef.Checked)
                _SelectTemplateType = SelectTemplateType.自訂;

            if (chkFileAllInOne.Checked)
                _SelectFileType = SelectFileType.單檔;

            if (chkFileSplitBySNum.Checked)
                _SelectFileType = SelectFileType.學號分檔;

            _Config.SetString("SelectTemplateType", _SelectTemplateType.ToString());
            _Config.SetString("SelectFileType", _SelectFileType.ToString());
            _Config.Save();        

            // 開始執行
            _bgWorker.RunWorkerAsync();
        }

        /// <summary>
        /// 新增 Data Table 內單筆欄位
        /// </summary>
        private void AddTableSColumn()
        {
            _dtTable.Columns.Add("學校名稱");
            _dtTable.Columns.Add("姓名");
            _dtTable.Columns.Add("性別");
            _dtTable.Columns.Add("學號");
            _dtTable.Columns.Add("入學年月");
            _dtTable.Columns.Add("入學學校");
            _dtTable.Columns.Add("畢業年月");
            _dtTable.Columns.Add("學籍核准文號");
            _dtTable.Columns.Add("畢業證書字號");
            _dtTable.Columns.Add("身分證號");
            _dtTable.Columns.Add("生日");
            _dtTable.Columns.Add("生日2");
            _dtTable.Columns.Add("出生地");
            _dtTable.Columns.Add("入學照片");
            _dtTable.Columns.Add("畢業照片");
            _dtTable.Columns.Add("入學照片2");
            _dtTable.Columns.Add("畢業照片2");
            _dtTable.Columns.Add("導師評語");
            _dtTable.Columns.Add("血型");
            _dtTable.Columns.Add("宗教");
            _dtTable.Columns.Add("生理狀態");
            _dtTable.Columns.Add("特殊疾病");
            _dtTable.Columns.Add("戶籍地址");
            _dtTable.Columns.Add("聯絡地址");
            _dtTable.Columns.Add("戶籍電話");
            _dtTable.Columns.Add("聯絡電話");
            _dtTable.Columns.Add("監護人_姓名");
            _dtTable.Columns.Add("監護人_關係");
            _dtTable.Columns.Add("監護人_通訊地址");
            _dtTable.Columns.Add("監護人_電話");
            _dtTable.Columns.Add("最喜歡的學科");
            _dtTable.Columns.Add("最感困難的學科");
            _dtTable.Columns.Add("特殊專長");
            _dtTable.Columns.Add("休閒興趣");
            _dtTable.Columns.Add("社團幹部");
            _dtTable.Columns.Add("班級幹部");
            _dtTable.Columns.Add("畢業成績");
            _dtTable.Columns.Add("最足以描述自己的幾句話");
            _dtTable.Columns.Add("最痛苦的回憶");
            _dtTable.Columns.Add("最快樂的回憶");
            _dtTable.Columns.Add("國中時的學校生活");
            _dtTable.Columns.Add("最不喜歡做的事_因為");
            _dtTable.Columns.Add("最不喜歡做的事");
            _dtTable.Columns.Add("最喜歡做的事_因為");
            _dtTable.Columns.Add("最喜歡做的事");
            _dtTable.Columns.Add("他是怎樣的人");
            _dtTable.Columns.Add("最要好的朋友");
            _dtTable.Columns.Add("喜歡的人_因為");
            _dtTable.Columns.Add("喜歡的人");
            _dtTable.Columns.Add("讀過且印象最深刻的課外書");
            _dtTable.Columns.Add("常指導我做功課的人");
            _dtTable.Columns.Add("家中最了解我的人");
            _dtTable.Columns.Add("升學意願");
            _dtTable.Columns.Add("就業意願");
            _dtTable.Columns.Add("職訓種類");
            _dtTable.Columns.Add("受訓地區");
            _dtTable.Columns.Add("職業意願");
            _dtTable.Columns.Add("就業地區");

            _dtTable.Columns.Add("父母關係");            
            _dtTable.Columns.Add("家庭氣氛");
            _dtTable.Columns.Add("父管教方式");
            _dtTable.Columns.Add("母管教方式");
            _dtTable.Columns.Add("居住環境");
            _dtTable.Columns.Add("本人住宿");
            _dtTable.Columns.Add("經濟狀況");
            _dtTable.Columns.Add("零用金");
            _dtTable.Columns.Add("零用金是否足夠");

            _dtTable.Columns.Add("自我認識_需要改進的地方_1");
            _dtTable.Columns.Add("自我認識_優點_1");
            _dtTable.Columns.Add("自我認識_個性_1");
            _dtTable.Columns.Add("自我認識_需要改進的地方_2");
            _dtTable.Columns.Add("自我認識_優點_2");
            _dtTable.Columns.Add("自我認識_個性_2");
            _dtTable.Columns.Add("自我認識_需要改進的地方_3");
            _dtTable.Columns.Add("自我認識_優點_3");
            _dtTable.Columns.Add("自我認識_個性_3");
            _dtTable.Columns.Add("生活感想_內容3_1");
            _dtTable.Columns.Add("生活感想_內容2_1");
            _dtTable.Columns.Add("生活感想_內容1_1");
            _dtTable.Columns.Add("生活感想_內容3_2");
            _dtTable.Columns.Add("生活感想_內容2_2");
            _dtTable.Columns.Add("生活感想_內容1_2");
            _dtTable.Columns.Add("備註_備註");

            _dtTable.Columns.Add("獎懲明細");

            //Cloud新增
            _dtTable.Columns.Add("原住民血統");
            _dtTable.Columns.Add("原住民血統_稱謂");
            _dtTable.Columns.Add("原住民血統_族別");
            _dtTable.Columns.Add("家中最了解我的人_因為");
            _dtTable.Columns.Add("我在家中最怕的人是");
            _dtTable.Columns.Add("我在家中最怕的人是_因為");
            _dtTable.Columns.Add("我覺得我的優點是");
            _dtTable.Columns.Add("我覺得我的缺點是");
            _dtTable.Columns.Add("最喜歡的國小（國中）老師");
            _dtTable.Columns.Add("最喜歡的國小（國中）老師__因為");
            _dtTable.Columns.Add("小學（國中）老師或同學常說我是");
            _dtTable.Columns.Add("小學（國中）時我曾在班上登任過的職務有");
            _dtTable.Columns.Add("我在小學（國中）得過的獎有");
            _dtTable.Columns.Add("我覺得我自己的過去最滿意的是");
            _dtTable.Columns.Add("我排遣休閒時間的方法是");
            _dtTable.Columns.Add("我最難忘的一件事是");
            _dtTable.Columns.Add("自我的心聲_一年級_我目前遇到最大的困難是");
            _dtTable.Columns.Add("自我的心聲_一年級_我目前最需要的協助是");
            _dtTable.Columns.Add("自我的心聲_二年級_我目前遇到最大的困難是");
            _dtTable.Columns.Add("自我的心聲_二年級_我目前最需要的協助是");
            _dtTable.Columns.Add("自我的心聲_三年級_我目前遇到最大的困難是");
            _dtTable.Columns.Add("自我的心聲_三年級_我目前最需要的協助是");
               

            
            // 動態新增
            for (int i = 1; i <= 300; i++)
            {
                _dtTable.Columns.Add("獎懲日期" + i);
                _dtTable.Columns.Add("獎懲類別支數" + i);
                _dtTable.Columns.Add("獎懲事由" + i);

                _dtTable.Columns.Add("獎勵日期" + i);
                _dtTable.Columns.Add("獎勵類別支數" + i);
                _dtTable.Columns.Add("獎勵事由" + i);
                _dtTable.Columns.Add("懲戒日期" + i);
                _dtTable.Columns.Add("懲戒類別支數" + i);
                _dtTable.Columns.Add("懲戒事由" + i);                
            }

            for (int i = 1; i <= _MaxDataCols; i++)
            {
                _dtTable.Columns.Add("異動日期" + i);
                _dtTable.Columns.Add("異動原因" + i);
                _dtTable.Columns.Add("缺曠統計" + i);
                _dtTable.Columns.Add("學年學期" + i);
                _dtTable.Columns.Add("學期成績" + i);
                _dtTable.Columns.Add("學習歷程_學年度" + i);
                _dtTable.Columns.Add("學習歷程_學期" + i);
                _dtTable.Columns.Add("學習歷程_班級" + i);
                _dtTable.Columns.Add("學習歷程_班導師" + i);
                _dtTable.Columns.Add("家長親屬_稱謂" + i);
                _dtTable.Columns.Add("家長親屬_姓名" + i);
                _dtTable.Columns.Add("家長親屬_出生年" + i);
                _dtTable.Columns.Add("家長親屬_存歿" + i);
                _dtTable.Columns.Add("家長親屬_教育程度" + i);
                _dtTable.Columns.Add("家長親屬_職業" + i);
                _dtTable.Columns.Add("家長親屬_工作機構" + i);
                _dtTable.Columns.Add("家長親屬_職稱" + i);
                _dtTable.Columns.Add("家長親屬_原國籍" + i);
                _dtTable.Columns.Add("兄弟姊妹_稱謂" + i);
                _dtTable.Columns.Add("兄弟姊妹_姓名" + i);
                _dtTable.Columns.Add("兄弟姊妹_出生年" + i);
                _dtTable.Columns.Add("兄弟姊妹_畢肄業學校" + i);

                _dtTable.Columns.Add("身高" + i);
                _dtTable.Columns.Add("體重" + i);
                _dtTable.Columns.Add("生活習慣" + i);
                _dtTable.Columns.Add("人際關係" + i);
                _dtTable.Columns.Add("外向行為" + i);
                _dtTable.Columns.Add("內向行為" + i);
                _dtTable.Columns.Add("學習動機" + i);
                _dtTable.Columns.Add("服務熱忱" + i);
                _dtTable.Columns.Add("人生態度" + i);
            }

            // 輔導相關測驗、晤談、個案、優先
            for (int i = 1; i <= 100; i++)
            {
                _dtTable.Columns.Add("測驗名稱" + i);
                _dtTable.Columns.Add("測驗日期" + i);
                _dtTable.Columns.Add("測驗結果" + i);
                _dtTable.Columns.Add("晤談紀錄日期" + i);
                _dtTable.Columns.Add("晤談紀錄對象" + i);
                _dtTable.Columns.Add("晤談紀錄方式" + i);
                _dtTable.Columns.Add("晤談紀錄內容要點" + i);
                _dtTable.Columns.Add("晤談紀錄記錄者姓名" + i);
                _dtTable.Columns.Add("個案會議會議日期" + i);
                _dtTable.Columns.Add("個案會議會議事由" + i);
                _dtTable.Columns.Add("個案會議內容要點" + i);
                _dtTable.Columns.Add("個案會議記錄者姓名" + i);
                _dtTable.Columns.Add("優先關懷立案日期" + i);
                _dtTable.Columns.Add("優先關懷個案類別" + i);
                _dtTable.Columns.Add("優先關懷個案來源" + i);
                _dtTable.Columns.Add("優先關懷記錄者姓名" + i);

            }
        }


        private void lnkDefaultView_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if(chkDefault.Checked)
                DownloadDefaultTemplate();
        }

        /// <summary>
        /// 下載預設樣板
        /// </summary>
        private void DownloadDefaultTemplate()
        {
            Document DefaultDoc = new Document(new MemoryStream(Properties.Resources.輔導資料紀錄表範本));

            if (DefaultDoc == null)
            {
                FISCA.Presentation.Controls.MsgBox.Show("預設範本發生錯誤無法產生.");
                return;
            }

            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "Word (*.doc)|*.doc";
            saveDialog.FileName = "輔導資料紀錄表範本";
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    DefaultDoc.Save(saveDialog.FileName);
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("儲存失敗。" + ex.Message);
                    return;
                }

                try
                {
                    System.Diagnostics.Process.Start(saveDialog.FileName);
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("開啟失敗。" + ex.Message);
                    return;
                }
            }        
        }

        private void lnkUserDef_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if(chkUserDef.Checked)
                DownloadUserDefTemplate();
        }

        /// <summary>
        /// 下載使用者自訂範本
        /// </summary>
        private void DownloadUserDefTemplate()
        {
            if (_Config.Template == null)
            {
                FISCA.Presentation.Controls.MsgBox.Show("目前沒有任何範本");
                return;
            }

            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "Word (*.doc)|*.doc";
            saveDialog.FileName = "輔導資料紀錄表自訂範本";
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {  
                     //   _Config.Template.ToDocument().Save(saveDialog.FileName); 改寫因 Aspose 新舊版相容問題
                    Document saDoc = new Document(new MemoryStream(_Config.Template.ToBinary()));
                    saDoc.Save(saveDialog.FileName);
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("儲存失敗。" + ex.Message);
                    return;
                }

                try
                {
                    System.Diagnostics.Process.Start(saveDialog.FileName);
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("開啟失敗。" + ex.Message);
                    return;
                }
            }
        }

        private void lnkUserDefUpload_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if(chkUserDef.Checked)
                UploadUserDefTemplate();
        }

        /// <summary>
        /// 上傳使用者自定範本
        /// </summary>
        private void UploadUserDefTemplate()
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Word (*.doc)|*.doc";
            if (openDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileInfo fileInfo = new FileInfo(openDialog.FileName);
                    TemplateType type = TemplateType.Word;
                    ReportTemplate template = new ReportTemplate(fileInfo, type);
                    _Config.Template = template;
                    _Config.Save();
                    FISCA.Presentation.Controls.MsgBox.Show("上傳範本成功");
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("上傳範本失敗" + ex.Message);
                }
            }        
        }

        private void lnkChkMappingField_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "Word (*.doc)|*.doc";
            saveDialog.FileName = "綜合表現紀錄表合併欄位說明";
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    Document doc = new Document(new MemoryStream(Properties.Resources.綜合表現紀錄表合併欄位說明));
                    doc.Save(saveDialog.FileName);
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("儲存失敗。" + ex.Message);
                    return;
                }

                try
                {
                    System.Diagnostics.Process.Start(saveDialog.FileName);
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("開啟失敗。" + ex.Message);
                    return;
                }
            }
        }
    }
}
