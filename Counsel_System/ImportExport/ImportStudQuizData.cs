using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Campus.DocumentValidator;
using Campus.Import;
using FISCA.UDT;
using K12.Data;
using System.Xml.Linq;

namespace Counsel_System2.ImportExport
{
    /// <summary>
    /// 匯入學生測驗資料
    /// </summary>
    public class ImportStudQuizData : ImportWizard
    {
        private ImportOption mOption;        
        DAO.UDTTransfer _UDTTransfer;        
        List<string> _ImportStudIDList;
        List<string> _DataFieldNameList;
        //XElement _elmRoot;
        string _QuizName;

        // 測驗ID
        string _QuizID;
        //private StringBuilder mstrLog = new StringBuilder();

        /// <summary>
        /// 設定測驗編號
        /// </summary>
        /// <param name="QuizID"></param>
        public void SetQuizID(string QuizID)
        {
            _QuizID = QuizID;
        
        }

        public ImportStudQuizData()
        {
            this.IsSplit = false;
            this.IsLog = false;
        }


        /// <summary>
        /// 設定測驗名稱
        /// </summary>
        /// <param name="Name"></param>
        public void SetQuizName(string Name)
        {
            _QuizName = Name;
        }
        /// <summary>
        /// 設定欄位
        /// </summary>
        /// <param name="DataFieldList"></param>
        public void SetDataFieldNameList(List<string> DataFieldNameList)
        {
            _DataFieldNameList = DataFieldNameList;
        }

 
        /// <summary>
        /// 取得支援的匯入動作
        /// </summary>
        /// <returns></returns>
        public override ImportAction GetSupportActions()
        {
            return ImportAction.InsertOrUpdate | ImportAction.Delete;
        }

        /// <summary>
        /// 分批執行匯入
        /// </summary>
        /// <param name="Rows">IRowStream物件列表</param>
        /// <returns>分批匯入完成訊息</returns>
        public override string Import(List<IRowStream> Rows)
        {        
          
            
            if (mOption.Action == ImportAction.InsertOrUpdate)
            {
                try
                {
                    // log
                    DAO.LogTransfer _logTransfer = new DAO.LogTransfer();

                    List<DAO.UDT_StudQuizDataDef> StudQuizDataList = new List<DAO.UDT_StudQuizDataDef>();
                    List<DAO.UDT_StudQuizDataDef> DelQuizDataList = new List<DAO.UDT_StudQuizDataDef>();
                    List<DAO.UDT_StudQuizDataDef> HasQuizDataList = new List<DAO.UDT_StudQuizDataDef>();
                    // 新增 Log
                    Dictionary<int,StringBuilder> Log_Insert =new Dictionary<int,StringBuilder> ();
                    // 刪除 Log
                    Dictionary<int, StringBuilder> Log_Delete = new Dictionary<int, StringBuilder>();


                    List<string> studentNumberList = new List<string>();
                    List<string> ClassSeatNoList = new List<string>();
                    
                    // 取得試別有學生資料
                    HasQuizDataList = _UDTTransfer.GetStudQuizDataByQuizID(_QuizID);
                    int count = 0;
                    string AuthorID = Utility.GetAuthorID();
                    foreach (IRowStream irs in Rows)
                    {
                        count++;
                        this.ImportProgress = count;
                        int sid = 0;
                        // 依學號比對
                        if (irs.Contains("學號") && irs.Contains("狀態"))
                        {
                            if (Global._StudentStatusDBDict.ContainsKey(irs.GetValue("狀態")))
                            {
                                sid = Utility.GetStudentID(irs.GetValue("學號"), irs.GetValue("狀態"));
                                studentNumberList.Add("學號:"+irs.GetValue("學號"));
                            }
                        }

                        // 依班座比對
                        if (irs.Contains("班級") && irs.Contains("座號") && irs.Contains("狀態"))
                        {
                            if (Global._StudentStatusDBDict.ContainsKey(irs.GetValue("狀態")))
                            {
                                sid = Utility.GetStudentID(irs.GetValue("班級"), irs.GetValue("座號"), irs.GetValue("狀態"));
                                ClassSeatNoList.Add("班級:" + irs.GetValue("班級") + ",座號:" + irs.GetValue("座號"));
                            }
                        }

                            // 比對需要刪除
                            foreach(DAO.UDT_StudQuizDataDef data in HasQuizDataList.Where(x=>x.StudentID ==sid))
                            {
                                DelQuizDataList.Add(data);
                            }

                         

                            DAO.UDT_StudQuizDataDef sqd = new DAO.UDT_StudQuizDataDef();
                            sqd.QuizID = int.Parse(_QuizID);
                            sqd.AuthorID = AuthorID;
                            sqd.StudentID = sid;

                            DateTime dt1, dt2;
                            if (irs.Contains("實施日期"))
                                if (DateTime.TryParse(irs.GetValue("實施日期"), out dt1))
                                    sqd.ImplementationDate = dt1;

                            if (irs.Contains("解析日期"))
                                if (DateTime.TryParse(irs.GetValue("解析日期"), out dt2))
                                    sqd.AnalysisDate = dt2;

                            List<XElement> elmList = new List<XElement>();
                            foreach (string str in _DataFieldNameList)
                            {
                                if (irs.Contains(str))
                                {
                                    XElement elm = new XElement("Item");
                                    elm.SetAttributeValue("name", str);
                                    elm.SetAttributeValue("value", irs.GetValue(str));
                                    elmList.Add(elm);
                                }
                            }
                            sqd.Content = Utility.ConvertXmlListToString1(elmList);

                            StudQuizDataList.Add(sqd);
                        
                    }

                    if (DelQuizDataList.Count > 0)
                    {
                        List<int> intList = (from data in DelQuizDataList select data.StudentID).ToList();
                        Dictionary<int, string> studNameDict = Utility.GetConvertStringDict1fromDB(intList);

                        // 收集 Log
                        foreach (DAO.UDT_StudQuizDataDef data in DelQuizDataList)
                        {
                            if(studNameDict.ContainsKey(data.StudentID))
                            {
                                StringBuilder sb = new StringBuilder();
                                sb.AppendLine("匯入輔導測驗-刪除");
                                sb.AppendLine(studNameDict[data.StudentID]);
                                sb.AppendLine("測驗名稱：" + _QuizName);
                                if (data.ImplementationDate.HasValue)
                                    sb.AppendLine("實施日期：" + data.ImplementationDate.Value.ToShortDateString());

                                if (data.AnalysisDate.HasValue)
                                    sb.AppendLine("解析日期：" + data.AnalysisDate.Value.ToShortDateString());

                                XElement elmContent = Utility.ConvertStringToXelm1(data.Content);
                                if (elmContent != null)
                                {
                                    foreach (XElement elm in elmContent.Elements("Item"))
                                    {
                                        if (elm.Attribute("name") != null && elm.Attribute("value") != null)
                                            sb.AppendLine("項目名稱：" + elm.Attribute("name").Value + " , 測驗結果：" + elm.Attribute("value").Value);
                                    }
                                }

                                if (!Log_Delete.ContainsKey(data.StudentID))
                                    Log_Delete.Add(data.StudentID, sb);                            
                            }
                        }

                        _UDTTransfer.DeleteStudQuizDataLlist(DelQuizDataList);
                    }
                    if (StudQuizDataList.Count > 0)
                    {
                        List<int> intList = (from data in StudQuizDataList select data.StudentID).ToList();
                        Dictionary<int, string> studNameDict = Utility.GetConvertStringDict1fromDB(intList);

                        // 收集 Log
                        foreach (DAO.UDT_StudQuizDataDef data in StudQuizDataList)
                        {
                            if (studNameDict.ContainsKey(data.StudentID))
                            {
                                StringBuilder sb = new StringBuilder();
                                // 檢查是否有刪除，有放上面
                                if (Log_Delete.ContainsKey(data.StudentID))
                                {
                                    sb.AppendLine(Log_Delete[data.StudentID].ToString());
                                    sb.AppendLine();
                                }

                                sb.AppendLine("匯入輔導測驗-新增");
                                sb.AppendLine(studNameDict[data.StudentID]);
                                sb.AppendLine("測驗名稱：" + _QuizName);
                                if (data.ImplementationDate.HasValue)
                                    sb.AppendLine("實施日期：" + data.ImplementationDate.Value.ToShortDateString());

                                if (data.AnalysisDate.HasValue)
                                    sb.AppendLine("解析日期：" + data.AnalysisDate.Value.ToShortDateString());

                                XElement elmContent = Utility.ConvertStringToXelm1(data.Content);
                                if (elmContent != null)
                                {
                                    foreach (XElement elm in elmContent.Elements("Item"))
                                    {
                                        if (elm.Attribute("name") != null && elm.Attribute("value") != null)
                                            sb.AppendLine("項目名稱：" + elm.Attribute("name").Value + " , 測驗結果：" + elm.Attribute("value").Value);
                                    }
                                }

                                if (!Log_Insert.ContainsKey(data.StudentID))
                                    Log_Insert.Add(data.StudentID, sb);
                            }
                        }

                        _UDTTransfer.InsertStudQuizDataLlist(StudQuizDataList);
                    }

                    // log 資料
                    if (Log_Insert.Count > 0)
                    {
                        foreach (KeyValuePair<int, StringBuilder> data in Log_Insert)
                            _logTransfer.SaveLog("輔導系統-匯入測驗資料", "匯入", "student", data.Key.ToString(), data.Value);
                    }

                    // 總共
                    StringBuilder logData = new StringBuilder();
                    logData.AppendLine("測驗名稱：" + _QuizName);
                    if (ClassSeatNoList.Count > 0)
                    {
                        logData.AppendLine("--依班級座號匯入--");
                        foreach (string str in ClassSeatNoList)
                            logData.AppendLine(str);

                        logData.AppendLine("匯入學生共"+ClassSeatNoList.Count+"人");
                    }
                    if (studentNumberList.Count > 0)
                    {
                        logData.AppendLine("--依學號匯入--");
                        foreach (string str in studentNumberList)
                            logData.AppendLine(str);

                        logData.AppendLine("匯入學生共" + studentNumberList.Count + "人");
                    }
                    // 寫入 log
                    _logTransfer.SaveLog("輔導系統.匯入測驗資料", "匯入", "", "", logData);
                  
                }
                catch(Exception ex)
                {
                    throw ex;
                }              
            }
            return "";            
        }

        public override string GetValidateRule()
        {
            if (Global._ImportStudQuizDataValElement != null)
            {
                StringBuilder sb = new StringBuilder();                

                string str="<?xml version=_1.0_ encoding=_utf-8_ ?>";
                string str1="<?xml-stylesheet type=_text/xsl_ href=_format.xsl_ ?>";
                sb.AppendLine(str.Replace('_','"'));
                sb.AppendLine(str1.Replace('_', '"'));
                sb.Append(Global._ImportStudQuizDataValElement.ToString());
                return sb.ToString();
            }
            else
                return Properties.Resources.ImportStudQuizDataVal_SNum;
        }

        public override void Prepare(ImportOption Option)
        {
            mOption = Option;            
            _UDTTransfer = new DAO.UDTTransfer();
            //_StudRecList = Student.SelectAll();
            _ImportStudIDList = new List<string>();
   

        }
    }


}
