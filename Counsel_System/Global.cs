using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Xml;
using K12.Data;

namespace Counsel_System2
{  
    class Global
    {
        /// <summary>
        /// 當有錯誤訊息
        /// </summary>
        public static StringBuilder _ErrorMessageList = new StringBuilder();

        /// <summary>
        /// 設定 ConfigData 名稱
        /// </summary>
        public static string CounselUserDefineDataRootName = "Counsel.UserConfigData";

        /// <summary>
        /// 設定 ConfigData 名稱
        /// </summary>
        public static string CounselUserDefineDataName = "UserConfigData";

        /// <summary>
        /// 學生測驗資料用的暫存
        /// </summary>
        public static Dictionary<string, string> _HasStudQuizDataDictTemp = new Dictionary<string, string>();
        /// <summary>
        /// 學生測驗資料XML
        /// </summary>
        public static XElement _ImportStudQuizDataValElement;

        /// <summary>
        /// 取得所有教師名稱對應用的暫存
        /// </summary>
        public static Dictionary<string, int> _AllTeacherNameIdDictTemp = new Dictionary<string, int>();

        /// <summary>
        /// 所有學生學號與狀態對應的暫存
        /// </summary>
        public static Dictionary<string, int> _AllStudentNumberStatusIDTemp = new Dictionary<string, int>();

        /// <summary>
        /// 學生狀態資料庫內的對應
        /// </summary>
        public static Dictionary<string, string> _StudentStatusDBDict = new Dictionary<string, string>();


        /// <summary>
        /// 設定輔導自訂欄位畫面選項
        /// </summary>
        public static Dictionary<string, string> _CounselUserDefSelectItemList = new Dictionary<string, string>();

        /// <summary>
        /// 啟動預設值
        /// </summary>
        public static void StartDefaultValue()
        {
            _CounselUserDefSelectItemList.Add("文字", "String");
            _CounselUserDefSelectItemList.Add("數字", "Number");
            _CounselUserDefSelectItemList.Add("日期", "Date");
        }

        /// <summary>
        /// XMLString轉成 Dictionary<string, string>
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public static Dictionary<string, string> CounselXMLToDictP1(string XmlDataStr)
        {
            // 如果重複FieldName 重複跳過
            Dictionary<string, string> retValue = new Dictionary<string, string>();

            if (!string.IsNullOrEmpty(XmlDataStr))
            {
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(XmlDataStr);

                XmlElement elms = doc.SelectSingleNode(CounselUserDefineDataName) as XmlElement;

                if (elms != null)
                    foreach (XmlElement xe in elms)
                        if (!retValue.ContainsKey(xe.GetAttribute("FieldName")))
                            retValue.Add(xe.GetAttribute("FieldName"), xe.GetAttribute("FieldType"));
            }
            return retValue;
        }


        /// <summary>
        /// 取得系統內設定自訂欄位資料項目與型態
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, string> GetUserConfigData()
        {               
            K12.Data.Configuration.ConfigData cd = K12.Data.School.Configuration[CounselUserDefineDataRootName];
            return CounselXMLToDictP1(cd[CounselUserDefineDataName]);
        }


        /// <summary>
        /// Dictionary<string, string>轉成 XML
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public static string CounselXMLToDictP1(Dictionary<string, string> data)
        {
            XmlElement elm = new XmlDocument().CreateElement(CounselUserDefineDataName);
            foreach (KeyValuePair<string, string> value in data)
            {
                XmlElement elmData = elm.OwnerDocument.CreateElement("Data");
                elmData.SetAttribute("FieldName", value.Key);
                elmData.SetAttribute("FieldType", value.Value);
                elm.AppendChild(elmData);
            }
            return elm.OuterXml;
        }

        /// <summary>
        /// AB表樣板交換者
        /// </summary>
        public static DAO.ABCardTemplateTransfer _ABCardTemplateTransfer = new DAO.ABCardTemplateTransfer();

        /// <summary>
        /// AB表存放學生學期歷程
        /// </summary>
        public static Dictionary<string, List<SemesterHistoryItem>> _StudentSemesterHistoryItemDict = new Dictionary<string, List<SemesterHistoryItem>>();

        /// <summary>
        /// AB表存放學生日常生活表現具體建議
        /// </summary>
        public static Dictionary<string,List<DAO.ABCard_StudentTextScore>> _ABCard_StudentTextScoreDict= new Dictionary<string,List<DAO.ABCard_StudentTextScore>> ();

        /// <summary>
        /// AB表存放學生畢業資訊
        /// </summary>
        public static Dictionary<string, LeaveInfoRecord> _ABCard_StudentGraduateDict = new Dictionary<string, LeaveInfoRecord>();

        /// <summary>
        /// AB表存放學生異動資料
        /// </summary>
        public static Dictionary<string, List<DAO.AB_StudDateString>> _ABCard_StudentUpdateRecDict = new Dictionary<string, List<DAO.AB_StudDateString>>();

        /// <summary>
        /// AB表存放獎懲資料
        /// </summary>
        public static Dictionary<string,List<DAO.AB_StudDateString>> _ABCard_StudentMDRecordDict = new Dictionary<string,List<DAO.AB_StudDateString>> ();

        /// <summary>
        /// AB表存放學生社團
        /// </summary>
        public static Dictionary<string, List<DAO.AB_StudText>> _ABCard_StudentSpecCourseDict = new Dictionary<string, List<DAO.AB_StudText>>();

        /// <summary>
        /// AB表存放學生擔任幹部
        /// </summary>
        public static Dictionary<string, List<DAO.AB_StudText>> _ABCard_StudentTheCadreDict = new Dictionary<string, List<DAO.AB_StudText>>();

        /// <summary>
        /// AB表存放學生學期領域成績與畢業成績
        /// </summary>
        public static Dictionary<string, DAO.AB_StudSemsDomainScore> _AB_StudSemsDomainScoreDict = new Dictionary<string, DAO.AB_StudSemsDomainScore>();

        /// <summary>
        /// AB 表存放輔導個案會議
        /// </summary>
        public static Dictionary<string, List<DAO.AB_RptCounselData>> _AB_CaseMeetingRecordToABRptDataDict = new Dictionary<string, List<DAO.AB_RptCounselData>>();

        /// <summary>
        /// AB 表存放輔導優先關懷
        /// </summary>
        public static Dictionary<string, List<DAO.AB_RptCounselData>> _AB_CareRecordToABRptDataDict = new Dictionary<string, List<DAO.AB_RptCounselData>>();

        /// <summary>
        /// AB 表存放輔導晤談紀錄
        /// </summary>
        public static Dictionary<string, List<DAO.AB_RptCounselData>> _AB_InterViewDataToABRptDataDict = new Dictionary<string, List<DAO.AB_RptCounselData>>();

        /// <summary>
        /// AB 表存放學生入學照片
        /// </summary>
        public static Dictionary<string, string> _AB_StudentFreshmanDict = new Dictionary<string, string>();

        /// <summary>
        /// AB 表存放畢業照片
        /// </summary>
        public static Dictionary<string, string> _AB_StudentGraduateDict = new Dictionary<string, string>();

        public static List<string> _StudentNumberList = new List<string>();

        // 存放測驗匯入班級座號檢查用
        public static Dictionary<string, int> _AllStudentClassSeatNoDictTemp = new Dictionary<string, int>();

        /// <summary>
        /// 優先關懷項目管理
        /// </summary>
        public static DAO.CareRecordItemManager _CareRecordItemManager = new DAO.CareRecordItemManager();
    }
}
