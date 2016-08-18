using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Data;
using Counsel_System2.DAO;
using System.Xml.Linq;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using Aspose.Cells;
using System.IO;
using K12.Data;

namespace Counsel_System2.ImportExport
{
    /// <summary>
    /// 匯出輔導綜合紀錄表
    /// </summary>
    public class ExportABCardData
    {
        BackgroundWorker _bgWorker;
        /// <summary>
        /// 資料表名稱
        /// </summary>
        List<string> _tableNameList;
        /// <summary>
        /// 資料表欄位名稱
        /// </summary>
        Dictionary<string, List<string>> _dtColumnsDict;
        /// <summary>
        /// 資料表內容
        /// </summary>
        Dictionary<string, DataTable> _dtTablesDict;
        /// <summary>
        /// 綜合紀錄表內容資料
        /// </summary>
        Dictionary<string,Dictionary<string,XElement>> _ABCardDataDict;

        /// <summary>
        /// 學生基本資料
        /// </summary>
        Dictionary<string, StudentRecord> _studRecDict;

        /// <summary>
        /// 暫存學生綜合紀錄資料使用 StudentID,TableName,ColumnName,Value
        /// </summary>
        Dictionary<string,Dictionary<string,List<Dictionary<string,string>>>> _StudABCardDataTempDict;

        /// <summary>
        /// 綜合紀錄表答案
        /// </summary>
        List<UDT_ABCardDataDef> _ABCardDataList;


        // 資料表內是固定單一值處理
        List<string> _tableName1List;


        List<string> _StudentIDList;

        public ExportABCardData(List<string> StudentIDList)
        {
            _StudentIDList = StudentIDList;
            _studRecDict = new Dictionary<string, StudentRecord>();
            foreach (StudentRecord stud in Student.SelectByIDs(_StudentIDList))
                _studRecDict.Add(stud.ID, stud);

            _tableName1List = new List<string>();
            _tableName1List.Add("本人概況");
            _tableName1List.Add("自傳");
            _tableName1List.Add("生活感想");
            _tableName1List.Add("畢業後計畫");
            _tableName1List.Add("適應情形");
            _tableName1List.Add("例外");
            _StudABCardDataTempDict = new Dictionary<string, Dictionary<string, List<Dictionary<string, string>>>>();
            _tableNameList = new List<string>();
            _dtColumnsDict = new Dictionary<string, List<string>>();
            _dtTablesDict = new Dictionary<string, DataTable>();           
            _ABCardDataDict = new Dictionary<string, Dictionary<string, XElement>>();
            _ABCardDataList = new List<UDT_ABCardDataDef>();
            _bgWorker = new BackgroundWorker();
            _bgWorker.DoWork += new DoWorkEventHandler(_bgWorker_DoWork);
            _bgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgWorker_RunWorkerCompleted);
        
        }

        void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            
            ExportDataTableToExcel();
        }

        void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            CreateDataTableAndColunms();
            FillDataToDataTable();
           
        }

        public void Go()
        {
            _bgWorker.RunWorkerAsync();
        }

        /// <summary>
        /// 資料存方資料相關 DataTable
        /// </summary>
        private void CreateDataTableAndColunms()
        {
            _tableNameList.Clear();
            _tableNameList.Add("本人概況");
            _tableNameList.Add("家庭狀況-監護人");
            _tableNameList.Add("家庭狀況-直系血親");
            _tableNameList.Add("家庭狀況-兄弟姊妹");
            _tableNameList.Add("家庭狀況-其它項目");
            _tableNameList.Add("學習狀況");
            _tableNameList.Add("自傳");
            _tableNameList.Add("自我認識");
            _tableNameList.Add("生活感想");
            _tableNameList.Add("畢業後計畫");
            _tableNameList.Add("適應情形");
            _tableNameList.Add("例外");

            _dtColumnsDict.Clear();
            _dtTablesDict.Clear();
            foreach (string name in _tableNameList)
            {
                List<string> columns = getColumnsName(name);
                _dtColumnsDict.Add(name, columns);
                
                DataTable dt = new DataTable();
                dt.TableName = name;
                foreach (string str in columns)
                    dt.Columns.Add(str);

                _dtTablesDict.Add(name, dt);            
            }           
        
        }

        /// <summary>
        /// 取的欄位名稱
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        private List<string> getColumnsName(string sheetName)
        {
            List<string> retVal = new List<string>();            

            switch (sheetName)
            {
                case "本人概況":
                    retVal = new string[] { "學號", "學生姓名", "血型", "宗教", "生理缺陷", "曾患特殊疾病", "身高一上", "身高一下", "身高二上", "身高二下", "身高三上", "身高三下", "體重一上", "體重一下", "體重二上", "體重二下", "體重三上", "體重三下", "備註","學生狀態" }.ToList();
                    break;
                case "家庭狀況-監護人":
                    retVal = new string[] { "學號", "學生姓名", "監護人姓名", "監護人性別", "監護人電話", "監護人關係", "監護人通訊地址","學生狀態" }.ToList();
                    break;
                case "家庭狀況-直系血親":
                    retVal = new string[] { "學號", "學生姓名", "稱謂", "姓名", "存歿", "出生年", "職業", "工作機構", "職稱", "教育程度", "學生狀態" }.ToList();
                    break;
                case "家庭狀況-兄弟姊妹":
                    retVal = new string[] { "學號", "學生姓名", "稱謂", "姓名", "畢肆業學校", "出生年次", "備註", "學生狀態" }.ToList();
                    break;
                case "家庭狀況-其它項目":
                    retVal = new string[] { "學號", "學生姓名", "項目年級", "父母關係", "家庭氣氛", "父親管教方式", "母親管教方式", "居住環境", "本人住宿", "經濟狀況", "每星期零用錢(元)", "我覺得是否足夠", "學生狀態" }.ToList();
                    break;
                case "學習狀況":
                    retVal = new string[] { "學號", "學生姓名", "項目年級", "特殊專長", "休閒興趣", "最喜歡的學科", "最感困難的學科", "社團幹部", "班級幹部", "學生狀態" }.ToList();
                    break;
                case "自傳":
                    retVal = new string[] { "學號", "學生姓名", "家中最了解我的人", "常指導我做功課的人", "讀過且印象最深刻的課外書", "喜歡的人", "喜歡的人_因為", "最要好的朋友", "最要好的朋友他是怎樣的人", "最喜歡做的事", "最喜歡做的事_因為", "最不喜歡做的事", "最不喜歡做的事_因為", "國中時的學校生活", "最快樂的回憶", "最痛苦的回憶", "最足以描述自己的幾句話", "學生狀態" }.ToList();
                    break;
                case "自我認識":
                    retVal = new string[] { "學號", "學生姓名", "項目年級", "個性", "優點", "需要改進地方", "填寫日期", "學生狀態" }.ToList();
                    break;
                case "生活感想":
                    retVal = new string[] { "學號", "學生姓名", "一年級期望", "一年級為達理想所需要的努力", "一年級期望師長給幫助", "一年級填寫日期", "二年級一年來的感想", "二年級今後努力的目標", "二年級期望師長給幫助", "二年級填寫日期", "學生狀態" }.ToList();                    
                    break;
                case "畢業後計畫":
                    retVal = new string[] { "學號", "學生姓名", "升學意願", "就業意願", "職業訓練", "受訓地區", "將來職業意願一", "將來職業意願二", "將來職業意願三", "就業地區意願一", "就業地區意願二", "就業地區意願三", "學生狀態" }.ToList();
                    break;
                case "適應情形":
                    retVal = new string[] { "學號", "學生姓名", "項目年級", "生活習慣", "人際關係", "外向行為", "內向行為", "學習動機", "服務熱忱", "人生態度", "學生狀態" }.ToList();
                    break;
                case "例外":
                    retVal = new string[] {  "學號", "學生姓名","家庭狀況電話", "家庭狀況備註","家庭狀況兄弟姊妹我排行第幾","學生狀態" }.ToList();
                    break;
            }
            
            return retVal;
        }

        /// <summary>
        /// 將資料寫入 DataTable
        /// </summary>
        private void FillDataToDataTable()
        {


            // 取得對照表
            XElement elmDataMapping = XElement.Parse(Properties.Resources.ABCardTransferDataOldMapping);

            // 讀取 UDT 綜合紀錄表資料
            _ABCardDataList = UDTTransfer.GetABCardDataListByStudentList(_StudentIDList);
            // 分學生
            _ABCardDataDict.Clear();
            foreach (UDT_ABCardDataDef data in _ABCardDataList)
            {
                string sid = data.StudentID.ToString();
                
                // 取得答案內容
                XElement elm=null;
                if (!string.IsNullOrWhiteSpace(data.Content))                
                    elm = XElement.Parse(data.Content);
                
                if (!_ABCardDataDict.ContainsKey(sid))
                {
                    Dictionary<string,XElement> elms = new Dictionary<string,XElement>();
                    elms.Add(data.SubjectName,elm);
                    _ABCardDataDict.Add(sid, elms);
                }

                if (!_ABCardDataDict[sid].ContainsKey(data.SubjectName))
                    _ABCardDataDict[sid].Add(data.SubjectName, elm);
            }

            try
            {

                
                // 讀取學生綜合紀錄表UDT內資料，並解析XML放在暫存 Dictionary內。
                // 清空暫存
                _StudABCardDataTempDict.Clear();
                foreach (KeyValuePair<string, Dictionary<string,XElement>> data in _ABCardDataDict)
                {
                    // StudentID
                    string studentId = data.Key;

                    if (!_StudABCardDataTempDict.ContainsKey(studentId))
                    {
                        Dictionary<string, List<Dictionary<string, string>>> tmpDict = new Dictionary<string, List<Dictionary<string, string>>>();
                        // tablename
                        foreach (string tbName in _tableNameList)
                        {
                            List<Dictionary<string, string>> tmpDictList = new List<Dictionary<string, string>>();
                            tmpDict.Add(tbName, tmpDictList);
                        }

                        _StudABCardDataTempDict.Add(studentId, tmpDict);
                    }

                    // 8 大項內容
                    
                    foreach (KeyValuePair<string, XElement> data1 in data.Value)
                    {                    
                        string GroupName = data1.Key;
                        XElement elm = data1.Value;

                        if (GroupName == "2.家庭狀況")
                        {
                            // 處理直系血親
                            List<string> tmpItemListFM = new string[] { "稱謂", "姓名", "存歿", "出生年", "職業", "工作機構", "職稱", "教育程度" }.ToList();

                            // 處理兄弟姊妹
                            List<string> tmpItemListBS = new string[] { "稱謂", "姓名", "畢肆業學校", "出生年次", "備註" }.ToList();


                            // 處理監護人
                            List<string> tmpItemList01 = new string[] { "監護人姓名", "監護人性別", "監護人電話", "監護人關係", "監護人通訊地址" }.ToList();
                            Dictionary<string, string> tmpParentDict = new Dictionary<string, string>();
                            foreach (string name in tmpItemList01)
                                tmpParentDict.Add(name, "");


                            // 處理其它項目
                            Dictionary<string, string> tmpOtherDict1 = new Dictionary<string, string>();
                            Dictionary<string, string> tmpOtherDict2 = new Dictionary<string, string>();
                            Dictionary<string, string> tmpOtherDict3 = new Dictionary<string, string>();
                            Dictionary<string, string> tmpOtherDict4 = new Dictionary<string, string>();
                            List<string> tmpItemsList1 = new string[] { "項目年級", "父母關係", "家庭氣氛", "父親管教方式", "母親管教方式", "居住環境", "本人住宿", "經濟狀況", "每星期零用錢(元)", "我覺得是否足夠" }.ToList();
                            // 建立對照
                            foreach (string name in tmpItemsList1)
                            {
                                if (name == "項目年級")
                                {
                                    tmpOtherDict1.Add(name, "1");
                                    tmpOtherDict2.Add(name, "2");
                                    tmpOtherDict3.Add(name, "3");
                                    tmpOtherDict4.Add(name, "4");
                                }
                                else
                                {
                                    tmpOtherDict1.Add(name, "");
                                    tmpOtherDict2.Add(name, "");
                                    tmpOtherDict3.Add(name, "");
                                    tmpOtherDict4.Add(name, "");
                                }
                            }

                            List<Dictionary<string, string>> tmpParentFMDictList = new List<Dictionary<string, string>>();


                            // 處理直系親屬資料，先建立稱謂再比對
                            foreach (XElement elm21 in elm.Elements("Ans"))
                            {
                                string name = elm21.Attribute("name").Value;
                                if (name == "AAB10000001")
                                {
                                    foreach (XElement elms in elm21.Elements("Item"))
                                    {
                                        Dictionary<string, string> tmpParentFMDict = new Dictionary<string, string>();
                                        foreach (string name1 in tmpItemListFM)
                                            tmpParentFMDict.Add(name1, "");

                                        foreach (XElement elms1 in elms.Elements("Field"))
                                        { 
                                            string skey=elms1.Attribute("key").Value;
                                            string svalue=elms1.Attribute("value").Value;

                                            switch (skey)
                                            {
                                                case "稱謂": tmpParentFMDict["稱謂"] = svalue; break;
                                                case "存、歿": tmpParentFMDict["存歿"] = svalue; break;
                                                case "出生年": tmpParentFMDict["出生年"] = svalue; break;
                                            }
                                        }                                        
                                        tmpParentFMDictList.Add(tmpParentFMDict);    
                                    }                                    
                                }
                                // 先找到父與母存儲位置
                                int idx = 0,idxF=0,idxM=0;
                                foreach (Dictionary<string, string> dataDict in tmpParentFMDictList)
                                {
                                    foreach (KeyValuePair<string, string> da01 in dataDict)
                                    {
                                        if (da01.Key == "稱謂" && da01.Value == "父")
                                            idxF = idx;

                                        if(da01.Key == "稱謂" && da01.Value == "母")
                                            idxM=idx;
                                    }
                                    idx++;
                                }

                                // 父教育程度
                                if (name == "AAB10000002")
                                    tmpParentFMDictList[idxF]["教育程度"] = elm21.Attribute("value").Value;
                                    

                                // 母教育程度
                                if (name == "AAB10000003")
                                    tmpParentFMDictList[idxM]["教育程度"] = elm21.Attribute("value").Value;

                                // 家長其它
                                if (name == "AAB10000004")
                                {
                                    XElement elmF = null;
                                    XElement elmM = null;
                                    foreach (XElement elmsss in elm21.Elements("Item"))
                                    {
                                        foreach (XElement elmsss1 in elmsss.Elements("Field"))
                                        {                                            
                                            string sssname = elmsss1.Attribute("key").Value;
                                            string sssvalue= elmsss1.Attribute("value").Value;

                                            if (sssvalue == "父")
                                                elmF = elmsss;

                                            if (sssvalue == "母")
                                                elmM = elmsss;
                                        }
                                    }

                                    if(elmF !=null)
                                    foreach (XElement elmd in elmF.Elements("Field"))
                                    {
                                        string sssname = elmd.Attribute("key").Value;
                                        string sssvalue = elmd.Attribute("value").Value;
                                        switch (sssname)
                                        {
                                            case "姓名": tmpParentFMDictList[idxF]["姓名"] = sssvalue; break;
                                            case "職業": tmpParentFMDictList[idxF]["職業"] = sssvalue; break;
                                            case "工作機構": tmpParentFMDictList[idxF]["工作機構"] = sssvalue; break;
                                            case "職稱": tmpParentFMDictList[idxF]["職稱"] = sssvalue; break;
                                        }
                                    }

                                    if (elmM != null)
                                        foreach (XElement elmd in elmF.Elements("Field"))
                                        {
                                            string sssname = elmd.Attribute("key").Value;
                                            string sssvalue = elmd.Attribute("value").Value;
                                            switch (sssname)
                                            {
                                                case "姓名": tmpParentFMDictList[idxM]["姓名"] = sssvalue; break;
                                                case "職業": tmpParentFMDictList[idxM]["職業"] = sssvalue; break;
                                                case "工作機構": tmpParentFMDictList[idxM]["工作機構"] = sssvalue; break;
                                                case "職稱": tmpParentFMDictList[idxM]["職稱"] = sssvalue; break;
                                            }
                                        }                                
                                }
                            }
                            // 加入 DataTable
                            foreach (Dictionary<string, string> daP in tmpParentFMDictList)
                                _StudABCardDataTempDict[studentId]["家庭狀況-直系血親"].Add(daP);

                            foreach (XElement elm24 in elm.Elements("Ans"))
                            {
                                string value="";
                                if(elm24.Attribute("value")!=null)
                                    value= elm24.Attribute("value").Value;

                                switch (elm24.Attribute("name").Value)
                                {
                                    // 兄弟姊妹
                                    case "AAB10000011":
                                        if (elm24.Element("Item") != null)
                                        {

                                            foreach (XElement elms in elm24.Elements("Item"))
                                            {
                                                Dictionary<string, string> tmpParentBSDict = new Dictionary<string, string>();
                                                foreach (string name in tmpItemListBS)
                                                    tmpParentBSDict.Add(name, "");
                                                foreach (XElement elms1 in elms.Elements("Field"))
                                                {
                                                    string skey = elms1.Attribute("key").Value;
                                                    string svalue = elms1.Attribute("value").Value;

                                                    switch (skey)
                                                    {
                                                        case "稱謂": tmpParentBSDict["稱謂"] = svalue; break;
                                                        case "姓名": tmpParentBSDict["姓名"] = svalue; break;
                                                        case "畢(肆)業學校": tmpParentBSDict["畢肆業學校"] = svalue; break;
                                                        case "出生年次": tmpParentBSDict["出生年次"] = svalue; break;
                                                        case "備註": tmpParentBSDict["備註"] = svalue; break;
                                                    }
                                                }
                                                _StudABCardDataTempDict[studentId]["家庭狀況-兄弟姊妹"].Add(tmpParentBSDict);
                                            }
                                        
                                        }
                                        break;

                                    // 監護人
                                    case "AAB10000005": tmpParentDict["監護人姓名"] = value; break;
                                    case "AAB10000006": tmpParentDict["監護人性別"] = value; break;
                                    case "AAB10000009": tmpParentDict["監護人電話"] = value; break;
                                    case "AAB10000007": tmpParentDict["監護人關係"] = value; break;
                                    case "AAB10000008": tmpParentDict["監護人通訊地址"] = value; break;

                                    // 其它
                                    case "AAB10000012": tmpOtherDict1["父母關係"] = value; break;
                                    case "AAB10000013": tmpOtherDict2["父母關係"] = value; break;
                                    case "AAB10000014": tmpOtherDict3["父母關係"] = value; break;
                                    case "AAB10000015": tmpOtherDict4["父母關係"] = value; break;
                                    case "AAB10000020": tmpOtherDict1["父親管教方式"] = value; break;
                                    case "AAB10000021": tmpOtherDict2["父親管教方式"] = value; break;
                                    case "AAB10000022": tmpOtherDict3["父親管教方式"] = value; break;
                                    case "AAB10000023": tmpOtherDict4["父親管教方式"] = value; break;
                                    case "AAB10000032": tmpOtherDict1["本人住宿"] = value; break;
                                    case "AAB10000033": tmpOtherDict2["本人住宿"] = value; break;
                                    case "AAB10000034": tmpOtherDict3["本人住宿"] = value; break;
                                    case "AAB10000035": tmpOtherDict4["本人住宿"] = value; break;
                                    case "AAB10000024": tmpOtherDict1["母親管教方式"] = value; break;
                                    case "AAB10000025": tmpOtherDict2["母親管教方式"] = value; break;
                                    case "AAB10000026": tmpOtherDict3["母親管教方式"] = value; break;
                                    case "AAB10000027": tmpOtherDict4["母親管教方式"] = value; break;
                                    case "AAB10000044": tmpOtherDict1["我覺得是否足夠"] = value; break;
                                    case "AAB10000045": tmpOtherDict2["我覺得是否足夠"] = value; break;
                                    case "AAB10000046": tmpOtherDict3["我覺得是否足夠"] = value; break;
                                    case "AAB10000047": tmpOtherDict4["我覺得是否足夠"] = value; break;
                                    case "AAB10000040": tmpOtherDict1["每星期零用錢(元)"] = value; break;
                                    case "AAB10000041": tmpOtherDict2["每星期零用錢(元)"] = value; break;
                                    case "AAB10000042": tmpOtherDict3["每星期零用錢(元)"] = value; break;
                                    case "AAB10000043": tmpOtherDict4["每星期零用錢(元)"] = value; break;
                                    case "AAB10000028": tmpOtherDict1["居住環境"] = value; break;
                                    case "AAB10000029": tmpOtherDict2["居住環境"] = value; break;
                                    case "AAB10000030": tmpOtherDict3["居住環境"] = value; break;
                                    case "AAB10000031": tmpOtherDict4["居住環境"] = value; break;
                                    case "AAB10000016": tmpOtherDict1["家庭氣氛"] = value; break;
                                    case "AAB10000017": tmpOtherDict2["家庭氣氛"] = value; break;
                                    case "AAB10000018": tmpOtherDict3["家庭氣氛"] = value; break;
                                    case "AAB10000019": tmpOtherDict4["家庭氣氛"] = value; break;
                                    case "AAB10000036": tmpOtherDict1["經濟狀況"] = value; break;
                                    case "AAB10000037": tmpOtherDict2["經濟狀況"] = value; break;
                                    case "AAB10000038": tmpOtherDict3["經濟狀況"] = value; break;
                                    case "AAB10000039": tmpOtherDict4["經濟狀況"] = value; break;
                                }                            
                            }

                            // 加入 Datatable
                            
                            
                            _StudABCardDataTempDict[studentId]["家庭狀況-監護人"].Add(tmpParentDict);
                            _StudABCardDataTempDict[studentId]["家庭狀況-其它項目"].Add(tmpOtherDict1);
                            _StudABCardDataTempDict[studentId]["家庭狀況-其它項目"].Add(tmpOtherDict2);
                            _StudABCardDataTempDict[studentId]["家庭狀況-其它項目"].Add(tmpOtherDict3);
                            _StudABCardDataTempDict[studentId]["家庭狀況-其它項目"].Add(tmpOtherDict4);
                        }
                        else if (GroupName == "3.學習狀況")
                        {                            
                            // 年級,欄位名稱,value
                            Dictionary<string, Dictionary<string, string>> _grData3Dict = new Dictionary<string, Dictionary<string, string>>();
                            foreach (XElement elm3 in elm.Elements("Ans"))
                            {
                              
                                foreach (XElement elmMap in (from da in elmDataMapping.Elements("Item") where da.Attribute("Key").Value == elm3.Attribute("name").Value select da))
                                {
                                    string AType = elmMap.Attribute("AType").Value;
                                    string col3Name = elmMap.Attribute("ColumnName").Value;
                                    if (!_grData3Dict.ContainsKey(AType))
                                    {
                                        Dictionary<string, string> tmp3Dict = new Dictionary<string, string>();
                                        _grData3Dict.Add(AType, tmp3Dict);
                                    }

                                    string value3 = "";
                                    if (elm3.Elements("Item").Count() == 0)
                                    {   // 單選
                                        if(elm3.Attribute("value") !=null)
                                            value3 = elm3.Attribute("value").Value;
                                    }
                                    else
                                    { 
                                        // 多選
                                        List<string> tmp3List = new List<string>();
                                        foreach (XElement tmp3Elm in elm3.Elements("Item"))
                                            tmp3List.Add(tmp3Elm.Attribute("value").Value);
                                        if (tmp3List.Count > 0)
                                            value3 = string.Join(",", tmp3List.ToArray());
                                    }

                                    if (!_grData3Dict[AType].ContainsKey(col3Name))
                                        _grData3Dict[AType].Add(col3Name,value3);
                                }
                            }
                            
                            // 放入暫存 Dict
                            foreach (KeyValuePair<string, Dictionary<string, string>> data3 in _grData3Dict)
                            {
                                Dictionary<string, string> tmp3DataDict = new Dictionary<string, string>();
                                tmp3DataDict.Add("項目年級", data3.Key);

                                foreach (KeyValuePair<string, string> data31 in data3.Value)
                                    if(!tmp3DataDict.ContainsKey(data31.Key)) 
                                        tmp3DataDict.Add(data31.Key,data31.Value);
                                _StudABCardDataTempDict[studentId]["學習狀況"].Add(tmp3DataDict);
                            }
                        }
                        else if (GroupName == "5.自我認識")
                        {
                            string tbName5 = "自我認識";
                            foreach (XElement elm5 in elm.Element("Ans").Elements("Item"))
                            {
                                Dictionary<string, string> tmpDict5 = new Dictionary<string, string>();
                                foreach (XElement elm51 in elm5.Elements("Field"))
                                {
                                    foreach (XElement elmMap in (from da in elmDataMapping.Elements("Item") where da.Attribute("TableName").Value == tbName5 select da))
                                    {                                        
                                        string columnName5 = elmMap.Attribute("ColumnName").Value;
                                        if(elm51.Attribute("key").Value == elmMap.Attribute("QuestionName").Value)
                                        if (!tmpDict5.ContainsKey(columnName5))
                                            tmpDict5.Add(columnName5, elm51.Attribute("value").Value);                                        
                                    }
                                }
                                _StudABCardDataTempDict[studentId][tbName5].Add(tmpDict5);                              
                            }
                        }
                        else
                        {
                            // 解析原有綜合紀錄表
                            foreach (XElement elm1 in elm.Elements("Ans"))
                            {
                                // 學生綜合紀錄答案key
                                string name = elm1.Attribute("name").Value;
                                string columnName = "";

                                // 比對對照表名稱與答案名稱，找出相對答案
                                foreach (string tableName in _tableName1List)
                                {
                                    Dictionary<string, string> tmpData1 = new Dictionary<string, string>();

                                    foreach (XElement elmMap in (from da in elmDataMapping.Elements("Item") where da.Attribute("Key").Value == name && da.Attribute("TableName").Value == tableName select da))
                                    {                                        
                                        columnName = elmMap.Attribute("ColumnName").Value;
                                        if (!tmpData1.ContainsKey(columnName))
                                        {
                                            string value = "";
                                            if (elm1.Attribute("value") != null)
                                                value = elm1.Attribute("value").Value;
                                            tmpData1.Add(columnName, value);
                                        }
                                    }

                                    _StudABCardDataTempDict[studentId][tableName].Add(tmpData1);
                                }
                            } 
                           
                            // 這段在解析有 CheckBox 類，不同儲存格式
                            List<string> ansListAAA10000011 = new List<string>();
                            List<string> ansListAAA10000012 = new List<string>();
                            List<string> andListAAG10000001 = new List<string>();
                            List<string> andListAAG10000002 = new List<string>();
                            List<string> andListAAG10000003 = new List<string>();
                            List<string> andListAAG10000004 = new List<string>();
                            foreach (XElement elm1 in elm.Elements("Ans"))
                            {
                                string name = elm1.Attribute("name").Value;
                                
                                // 生理缺陷
                                if (name == "AAA10000011")
                                    ansListAAA10000011 = elmConvertCheckItemToList(elm1);

                                // 曾患特殊疾病
                                if (name == "AAA10000012")
                                    ansListAAA10000012 = elmConvertCheckItemToList(elm1);

                                // 升學意願
                                if (name == "AAG10000001")
                                    andListAAG10000001 = elmConvertCheckItemToList(elm1);
                                // 就業意願
                                if (name == "AAG10000002")
                                    andListAAG10000002 = elmConvertCheckItemToList(elm1);
                                // 職業訓練
                                if (name == "AAG10000003")
                                    andListAAG10000003 = elmConvertCheckItemToList(elm1);
                                // 受訓地區
                                if (name == "AAG10000004")
                                    andListAAG10000004 = elmConvertCheckItemToList(elm1);


                            }
                             foreach (Dictionary<string, string> dassDict in _StudABCardDataTempDict[studentId]["本人概況"])
                             {
                                 if (dassDict.ContainsKey("生理缺陷"))
                                 {
                                     if(ansListAAA10000011.Count>0)
                                         dassDict["生理缺陷"] = string.Join(",", ansListAAA10000011.ToArray());
                                 }
                                 if (dassDict.ContainsKey("曾患特殊疾病"))
                                 {
                                     if(ansListAAA10000012.Count>0)
                                         dassDict["曾患特殊疾病"] = string.Join(",", ansListAAA10000012.ToArray());
                                 }
                             }

                             foreach (Dictionary<string, string> dassDict in _StudABCardDataTempDict[studentId]["畢業後計畫"])
                             {
                                 if (dassDict.ContainsKey("升學意願"))
                                 {
                                     if (andListAAG10000001.Count > 0)
                                         dassDict["升學意願"] = string.Join(",", andListAAG10000001.ToArray());
                                 }

                                 if (dassDict.ContainsKey("就業意願"))
                                 {
                                     if (andListAAG10000002.Count > 0)
                                         dassDict["就業意願"] = string.Join(",", andListAAG10000002.ToArray());
                                 }

                                 if (dassDict.ContainsKey("職業訓練"))
                                 {
                                     if (andListAAG10000003.Count > 0)
                                         dassDict["職業訓練"] = string.Join(",", andListAAG10000003.ToArray());
                                 }

                                 if (dassDict.ContainsKey("受訓地區"))
                                 {
                                     if (andListAAG10000004.Count > 0)
                                         dassDict["受訓地區"] = string.Join(",", andListAAG10000004.ToArray());
                                 }
                             }
                        }
                    }
                }


                // 讀取暫存 Dictionary將綜合紀錄表資料放入 DataTable 內。
                foreach (KeyValuePair<string,Dictionary<string,List<Dictionary<string,string>>>> data in _StudABCardDataTempDict)
                {
                    // StudentID
                    string studentId = data.Key;
                    string StudentNumber = _studRecDict[studentId].StudentNumber;
                    string StudentName = _studRecDict[studentId].Name;
                    string StudentStatus = _studRecDict[studentId].Status.ToString();

                    foreach (KeyValuePair<string, List<Dictionary<string, string>>> data1 in data.Value)
                    {
                        string TableName = data1.Key;
                        if (_dtTablesDict.ContainsKey(TableName))
                        {

                            if (TableName == "自我認識" || TableName=="學習狀況" || TableName=="家庭狀況-其它項目" || TableName=="家庭狀況-兄弟姊妹" || TableName=="家庭狀況-直系血親")
                            {
                                foreach (Dictionary<string, string> data2 in data1.Value)
                                {
                                    DataRow dr = _dtTablesDict[TableName].NewRow();
                                    foreach (KeyValuePair<string, string> data3 in data2)
                                        dr[data3.Key] = data3.Value;

                                    dr["學號"] = StudentNumber;
                                    dr["學生姓名"] = StudentName;
                                    dr["學生狀態"] = StudentStatus;
                                    _dtTablesDict[TableName].Rows.Add(dr);
                                }
                            }
                            else
                            {

                                DataRow dr = _dtTablesDict[TableName].NewRow();
                                foreach (Dictionary<string, string> data2 in data1.Value)
                                {                                
                                    foreach (KeyValuePair<string, string> data3 in data2)
                                        dr[data3.Key] = data3.Value;
                                }
                                dr["學號"] = StudentNumber;
                                dr["學生姓名"] = StudentName;
                                dr["學生狀態"] = StudentStatus;
                                _dtTablesDict[TableName].Rows.Add(dr);
                            }
                        }
                    
                    }
                }
            }
            catch (Exception ex)
            { 

            }

        }

        /// <summary>
        /// 轉換答案內 checkBox 型態內容
        /// </summary>
        /// <param name="elmItems"></param>
        /// <returns></returns>
        private List<string> elmConvertCheckItemToList(XElement elmItems)
        {
            List<string> retVal = new List<string>();
            foreach (XElement elmss in elmItems.Elements("Item"))
            {
                if (elmss.Attribute("remark") != null)
                    retVal.Add(elmss.Attribute("value").Value + ":" + elmss.Attribute("remark").Value);
                else
                    retVal.Add(elmss.Attribute("value").Value);
            }
            return retVal;
        }

        /// <summary>
        /// 將 DataTable 內資料匯出至 Excel
        /// </summary>
        private void ExportDataTableToExcel()
        {
            Workbook wb = new Workbook();
            
            foreach (KeyValuePair<string, DataTable> data in _dtTablesDict)
            {
                int wstIdx=wb.Worksheets.Add();

                wb.Worksheets[wstIdx].Name = data.Key;
                wb.Worksheets[wstIdx].Cells.ImportDataTable(data.Value, true, "A1");
            }
            wb.Worksheets.RemoveAt(0);
            CompletedXls("輔導綜合紀錄表", wb);
        }

        private void CompletedXls(string inputReportName, Workbook inputXls)
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
                wb.Save(path, Aspose.Cells.FileFormatType.Excel2003);
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                SaveFileDialog sd = new SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".xls";
                sd.Filter = "XLS檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        wb.Save(sd.FileName, Aspose.Cells.FileFormatType.Excel2003);

                    }
                    catch
                    {
                        MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
        }

    }
}
