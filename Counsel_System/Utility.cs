using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using K12.Data;
using FISCA.Data;
using System.Data;
using JHSchool.Data;
using Aspose.Words;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using Counsel_System2.DAO;
using Aspose.Words.Tables;

namespace Counsel_System2
{
    public class Utility
    {
        /// <summary>
        /// 將字串轉成XElement
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static XElement ConvertStringToXelm1(string str)
        {
            StringBuilder sb1 = new StringBuilder();
            sb1.Append("<root>"); sb1.Append(str); sb1.Append("</root>");
            return XElement.Parse(sb1.ToString());        
        }

        /// <summary>
        /// 將 XElementList轉成 String
        /// </summary>
        /// <param name="xmlList"></param>
        /// <returns></returns>
        public static string ConvertXmlListToString1(List<XElement> xmlList)
        {
            StringBuilder sb1 = new StringBuilder();

            foreach (XElement elm in xmlList)
                sb1.Append(elm.ToString());

            return sb1.ToString();        
        }

        /// <summary>
        /// 取得登入帳號
        /// </summary>
        /// <returns></returns>
        public static string GetAuthorID()
        {
            return FISCA.Authentication.DSAServices.UserAccount;
        }

        /// <summary>
        /// 取得非刪除教師名稱,TeacherID,TeacherName
        /// </summary>
        /// <returns></returns>
        public static Dictionary<int, string> GetTeacherIDNameDict()
        {
            Dictionary<int, string> retVal = new Dictionary<int, string>();
            foreach (TeacherRecord TRec in Teacher.SelectAll())
            {
                if (TRec.Status == TeacherRecord.TeacherStatus.刪除)
                    continue;

                int tid = int.Parse(TRec.ID);
                if (!retVal.ContainsKey(tid))
                    retVal.Add(tid, TRec.Name + TRec.Nickname);
            }

            return retVal;
        }

        ///// <summary>
        ///// 取得非刪除教師名稱,TeacherName,TeacherID
        ///// </summary>
        ///// <returns></returns>
        //public static Dictionary<string, int> GetTeacherNameIDDict()
        //{
        //    Dictionary<string, int> retVal = new Dictionary<string, int>();
        //    List<TeacherRecord> TeacherList = (from data in Teacher.SelectAll() where data.Status == TeacherRecord.TeacherStatus.一般 orderby data.Name, data.Nickname select data).ToList();
        //    foreach (K12.Data.TeacherRecord TRec in TeacherList)
        //    {
        //        int tid = int.Parse(TRec.ID);
        //        string TName = TRec.Name + TRec.Nickname;
        //        if (!retVal.ContainsKey(TName))
        //            retVal.Add(TName, tid);
        //    }

        //    return retVal;
        //}


        /// <summary>
        /// 取得輔導相關老師名稱,TeacherID,TeacherName
        /// </summary>
        /// <returns></returns>
        public static Dictionary<int, string> GetCounselTeacherIDNameDict(string StudentID)
        {
            
            Dictionary<int, string> retVal = new Dictionary<int, string>();
            QueryHelper helper = new QueryHelper();
            string sqlStr = "select teacher.id,teacher.teacher_name,teacher.nickname from $counsel.student_list inner join tag_teacher on tag_teacher.id=$counsel.student_list.ref_teacher_tag_id inner join teacher on teacher.id=tag_teacher.ref_teacher_id where $counsel.student_list.ref_student_id="+StudentID;
            DataTable dt = helper.Select(sqlStr);

            foreach (DataRow dr in dt.Rows)
            {
                int id = int.Parse(dr[0].ToString());

                string name = "";
                if(string.IsNullOrEmpty (dr[2].ToString ()))
                    name =dr[1].ToString ();
                else
                    name=dr[1].ToString ()+"("+dr[2].ToString ()+")";

                if(!retVal.ContainsKey(id))
                    retVal.Add(id,name );
            }

            // 加入班導師
            string sqlStr1 = "select teacher.id,teacher.teacher_name,teacher.nickname from student inner join class on student.ref_class_id = class.id inner join teacher on class.ref_teacher_id=teacher.id where student.id=" + StudentID;
            DataTable dt1 = helper.Select(sqlStr1);

            foreach (DataRow dr in dt1.Rows)
            {
                int id = int.Parse(dr[0].ToString());

                string name = "";
                if (string.IsNullOrEmpty(dr[2].ToString()))
                    name = dr[1].ToString();
                else
                    name = dr[1].ToString() + "(" + dr[2].ToString() + ")";

                if (!retVal.ContainsKey(id))
                    retVal.Add(id, name);
            }

            return retVal;
           
        }


        /// <summary>
        /// 取得輔導相關老師名稱,TeacherName,TeacherID
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, int> GetCounselTeacherNameIDDict(string StudentID)
        {

            Dictionary<string, int> retVal = new Dictionary<string, int>();
            QueryHelper helper = new QueryHelper();
            string sqlStr = "select teacher.id,teacher.teacher_name,teacher.nickname from $counsel.student_list inner join tag_teacher on tag_teacher.id=$counsel.student_list.ref_teacher_tag_id inner join teacher on teacher.id=tag_teacher.ref_teacher_id where $counsel.student_list.ref_student_id=" + StudentID;
            DataTable dt = helper.Select(sqlStr);

            foreach (DataRow dr in dt.Rows)
            {
                int id = int.Parse(dr[0].ToString());

                string name = "";
                if (string.IsNullOrEmpty(dr[2].ToString()))
                    name = dr[1].ToString();
                else
                    name = dr[1].ToString() + "(" + dr[2].ToString() + ")";

                if (!retVal.ContainsKey(name))
                    retVal.Add(name, id);
            }

            // 加入班導師
            string sqlStr1 = "select teacher.id,teacher.teacher_name,teacher.nickname from student inner join class on student.ref_class_id = class.id inner join teacher on class.ref_teacher_id=teacher.id where student.id=" + StudentID;
            DataTable dt1 = helper.Select(sqlStr1);

            foreach (DataRow dr in dt1.Rows)
            {
                int id = int.Parse(dr[0].ToString());

                string name = "";
                if (string.IsNullOrEmpty(dr[2].ToString()))
                    name = dr[1].ToString();
                else
                    name = dr[1].ToString() + "(" + dr[2].ToString() + ")";

                if (!retVal.ContainsKey(name))
                    retVal.Add(name, id);
            }
            return retVal;
        }

        /// <summary>
        /// 顯示必填欄位未填無法儲存訊息
        /// </summary>
        public static void ShowCannotSaveMessage()
        {
            FISCA.Presentation.Controls.MsgBox.Show("標示*必填欄位未填完整，請填寫後再儲存。");
        }

        /// <summary>
        /// 顯示儲存完成訊息
        /// </summary>
        public static void ShowSavedMessage()
        {
            FISCA.Presentation.Controls.MsgBox.Show("儲存完成。");
        }

        /// <summary>
        /// 透過輔導系統老師類別取得老師
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static List<DAO.CounselTeacherRecord> GetCounselTeacherByType(DAO.CounselTeacherRecord.CounselTeacherType type)
        { 
          
            List<DAO.CounselTeacherRecord> CounselTeacherList = new List<DAO.CounselTeacherRecord>();

            QueryHelper hepler = new QueryHelper();
            string strSQL = "select distinct teacher.id as id1,tag_teacher.id as id2,teacher.teacher_name,teacher.nickname,tag.name,tag.prefix from teacher inner join tag_teacher on teacher.id=tag_teacher.ref_teacher_id inner join tag on tag_teacher.ref_tag_id=tag.id where tag.prefix='輔導' and tag.name='"+type.ToString ()+"';";
            DataTable dt = hepler.Select(strSQL);


            foreach (DataRow dr in dt.Rows)
            { 
                DAO.CounselTeacherRecord counselTeacherRec = new DAO.CounselTeacherRecord ();
                // 教師編號
                counselTeacherRec.TeacherID = dr[0].ToString();
                // TeacherTagID
                counselTeacherRec.TeacherTag_ID=int.Parse(dr[1].ToString ());
                // 教師名稱
                if (string.IsNullOrEmpty(dr[3].ToString()))
                    counselTeacherRec.TeacherFullName = dr[2].ToString();
                else
                    counselTeacherRec.TeacherFullName = dr[2].ToString() + "(" + dr[3].ToString() + ")";

                // 類別
                counselTeacherRec.counselTeacherType = type;
                CounselTeacherList.Add(counselTeacherRec);
            }
            return CounselTeacherList.OrderBy(x=>x.TeacherFullName).ToList ();
        }
        

        /// <summary>
        /// 取得所有輔導相關老師
        /// </summary>
        /// <returns></returns>
        public static Dictionary<int, DAO.CounselTeacherRecord> GetAllCounselTeacherDict()
        {
            Dictionary<int, DAO.CounselTeacherRecord> CounselTeacherRecordDict = new Dictionary<int, DAO.CounselTeacherRecord>();            
            QueryHelper hepler = new QueryHelper ();
            string strSQL="select distinct teacher.id as id1,tag_teacher.id as id2,teacher.teacher_name,teacher.nickname,tag.name,tag.prefix from teacher inner join tag_teacher on teacher.id=tag_teacher.ref_teacher_id inner join tag on tag_teacher.ref_tag_id=tag.id where tag.prefix='輔導'";
            DataTable dt =hepler.Select (strSQL);

            foreach(DataRow dr in dt.Rows)
            {
                // TeachertagID
                int eid = int.Parse(dr[1].ToString ());
                if (!CounselTeacherRecordDict.ContainsKey(eid))
                {
                    DAO.CounselTeacherRecord cRec = new DAO.CounselTeacherRecord();
                    if (dr[4].ToString () == "輔導老師")
                        cRec.counselTeacherType = DAO.CounselTeacherRecord.CounselTeacherType.輔導老師;

                    if (dr[4].ToString() == "認輔老師")
                        cRec.counselTeacherType = DAO.CounselTeacherRecord.CounselTeacherType.認輔老師;

                    if (dr[4].ToString() == "輔導主任")
                        cRec.counselTeacherType = DAO.CounselTeacherRecord.CounselTeacherType.輔導主任;

                    cRec.TeacherTag_ID = eid;
                    cRec.TeacherID = dr[0].ToString();
                    if(string.IsNullOrEmpty(dr[3].ToString ()))
                        cRec.TeacherFullName = dr[2].ToString ();
                    else
                        cRec.TeacherFullName=dr[2].ToString ()+"("+dr[3].ToString ()+")";

                    CounselTeacherRecordDict.Add(eid, cRec);
                }
            }
            return CounselTeacherRecordDict;
        }

        /// <summary>
        /// 解析參與人員從XMLToStringA1,A2,A3
        /// </summary>
        /// <param name="XMLString"></param>
        /// <returns></returns>
        public static string AttendeesXMLToString(string XMLString)
        {            
            StringBuilder sb = new StringBuilder();
            sb.Append("<root>"); sb.Append(XMLString); sb.Append("</root>");
            // 解析 ContentXML
            XElement xmlAttendees = XElement.Parse(sb.ToString());
            List<string> itemList = new List<string>();
            if (xmlAttendees != null)
            {                
                // 參與人員 ---
                foreach (XElement elm in xmlAttendees.Elements("Item"))
                {
                    if (elm.Attribute("name") == null)
                        continue;

                    switch (elm.Attribute("name").Value.ToString())
                    {
                        // 學生
                        case "學生":
                            itemList.Add("學生");
                            break;
                        // 家長
                        case "家長":
                            itemList.Add("家長");
                            break;

                        // 專家
                        case "專家":
                            itemList.Add("專家");
                            break;

                        // 醫師
                        case "醫師":
                            itemList.Add("醫師");
                            break;

                        // 社工人員
                        case "社工人員":
                            itemList.Add("社工人員");
                            break;

                        // 導師
                        case "導師":
                            itemList.Add("導師");
                            break;

                        // 教官
                        case "教官":
                            itemList.Add("教官");
                            break;

                        // 輔導老師
                        case "輔導老師":
                            itemList.Add("輔導老師");
                            break;

                        // 任課老師
                        case "任課老師":
                            itemList.Add("任課老師");
                            break;

                        // 其它**
                        case "其它":                       
                            if (elm.Attribute("remark") != null)
                                itemList.Add("其它:" + elm.Attribute("remark").Value.ToString());
                            else
                                itemList.Add("其它:");
                            break;
                    }                   

                }
            }
            if (itemList.Count > 0)
                return string.Join(",", itemList.ToArray());
            else
                return string.Empty;
            
        }

        /// <summary>
        /// 解析輔導方式從XMLToStringA1,A2,A3
        /// </summary>
        /// <param name="XMLString"></param>
        /// <returns></returns>

        public static string CounselTypeXMLToString(string XMLString)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<root>"); sb.Append(XMLString); sb.Append("</root>");
            XElement xmlCounselType = XElement.Parse(sb.ToString());
            // 輔導方式--
            List<string> itemList = new List<string>();
            if (xmlCounselType != null)
            {
                foreach (XElement elm in xmlCounselType.Elements("Item"))
                {
                    if (elm.Attribute("name") == null)
                        continue;

                    switch (elm.Attribute("name").Value.ToString())
                    {
                        // 暫時結案
                        case "暫時結案":
                            itemList.Add("暫時結案");
                            break;

                        // 專案輔導
                        case "專案輔導":
                            itemList.Add("專案輔導");
                            break;
                        // 導師輔導
                        case "導師輔導":
                            itemList.Add("導師輔導");
                            break;
                        // 轉介**
                        case "轉介":                            
                            if (elm.Attribute("remark") != null)
                                itemList.Add("轉介:"+elm.Attribute("remark").Value.ToString());
                            else
                                itemList.Add("轉介:");
                            break;
                        // 就醫**
                        case "就醫":                            
                            if (elm.Attribute("remark") != null)
                                itemList.Add("就醫:"+elm.Attribute("remark").Value.ToString());
                            else
                                itemList.Add("就醫:");
                            break;
                        // 其它**
                        case "其它":                            
                            if (elm.Attribute("remark") != null)
                                itemList.Add("其它:"+elm.Attribute("remark").Value.ToString());
                            else
                                itemList.Add("其它:");
                            break;
                    }
                }
            }

            if (itemList.Count > 0)
                return string.Join(",", itemList.ToArray());
            else
                return string.Empty;
        }

        /// <summary>
        /// 解析輔導歸類從XMLToStringA1,A2,A3
        /// </summary>
        /// <param name="XMLString"></param>
        /// <returns></returns>
        public static string CounselTypeKindXMLToString(string XMLString)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<root>"); sb.Append(XMLString); sb.Append("</root>");
            XElement xmlCounselTypeKind = XElement.Parse(sb.ToString());
            List<string> itemList = new List<string>();
            // 輔導歸類--
            if (xmlCounselTypeKind != null)
            {
                foreach (XElement elm in xmlCounselTypeKind.Elements("Item"))
                {
                    if (elm.Attribute("name") == null)
                        continue;

                    switch (elm.Attribute("name").Value.ToString())
                    {
                        // 違規
                        case "違規":
                            itemList.Add("違規");
                            break;

                        // 遲曠
                        case "遲曠":
                            itemList.Add("遲曠");
                            break;

                        // 學習
                        case "學習":
                            itemList.Add("學習");
                            break;

                        // 生涯
                        case "生涯":
                            itemList.Add("生涯");
                            break;

                        // 人
                        case "人":
                            itemList.Add("人");
                            break;

                        // 人際
                        case "人際":
                            itemList.Add("人際");
                            break;

                        // 休退轉
                        case "休退轉":
                            itemList.Add("休退轉");
                            break;

                        // 家庭
                        case "家庭":
                            itemList.Add("家庭");
                            break;

                        // 師生
                        case "師生":
                            itemList.Add("師生");
                            break;

                        // 情感
                        case "情感":
                            itemList.Add("情感");
                            break;

                        // 精神
                        case "精神":
                            itemList.Add("精神");
                            break;

                        case "家暴": itemList.Add("家暴"); break;
                        case "霸凌": itemList.Add("霸凌"); break;
                        case "中輟": itemList.Add("中輟"); break;
                        case "性議題": itemList.Add("性議題"); break;
                        case "戒毒": itemList.Add("戒毒"); break;
                        case "網路成癮": itemList.Add("網路成癮"); break;
                        case "情緒障礙": itemList.Add("情緒障礙"); break;

                        // 其它**
                        case "其它":                          
                            if (elm.Attribute("remark") != null)
                                itemList.Add("其它:"+elm.Attribute("remark").Value.ToString());
                            else
                                itemList.Add("其它:");
                            break;

                    }
                }
            }
            if (itemList.Count > 0)
                return string.Join(",", itemList.ToArray());
            else
                return string.Empty;            
        }

        /// <summary>
        /// 解析參與人員成Dict
        /// </summary>
        /// <param name="XMLString"></param>
        /// <returns></returns>
        public static Dictionary<string, string> GetConvertCounselXMLVal_Attendees(string XMLString)
        { 
             StringBuilder sb = new StringBuilder();
            sb.Append("<root>"); sb.Append(XMLString); sb.Append("</root>");
            // 解析 ContentXML
            XElement xmlAttendees = XElement.Parse(sb.ToString());
            Dictionary<string,string> itemDict = new Dictionary<string,string> ();
            if (xmlAttendees != null)
            {                
                // 參與人員 ---
                foreach (XElement elm in xmlAttendees.Elements("Item"))
                {
                    if (elm.Attribute("name") == null)
                        continue;

                    switch (elm.Attribute("name").Value.ToString())
                    {
                        // 學生
                        case "學生":
                            itemDict.Add("參與人員:學生","1");
                            break;
                        // 家長
                        case "家長":
                            itemDict.Add("參與人員:家長","1");
                            break;

                        // 專家
                        case "專家":
                            itemDict.Add("參與人員:專家","1");
                            break;

                        // 醫師
                        case "醫師":
                            itemDict.Add("參與人員:醫師","1");
                            break;

                        // 社工人員
                        case "社工人員":
                            itemDict.Add("參與人員:社工人員","1");
                            break;

                        // 導師
                        case "導師":
                            itemDict.Add("參與人員:導師","1");
                            break;

                        // 教官
                        case "教官":
                            itemDict.Add("參與人員:教官","1");
                            break;

                        // 輔導老師
                        case "輔導老師":
                            itemDict.Add("參與人員:輔導老師","1");
                            break;

                        // 任課老師
                        case "任課老師":
                            itemDict.Add("參與人員:任課老師","1");
                            break;

                        // 其它**
                        case "其它":                       
                            if (elm.Attribute("remark") != null)
                                itemDict.Add("參與人員:其它備註",elm.Attribute("remark").Value.ToString());
                            
                                itemDict.Add("參與人員:其它","1");
                            break;
                    }                   

                }
            }
            return itemDict ;
        }

        /// <summary>
        /// 解析輔導方式成Dict
        /// </summary>
        /// <param name="XMLString"></param>
        /// <returns></returns>
        public static Dictionary<string, string> GetConvertCounselXMLVal_CounselType(string XMLString)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<root>"); sb.Append(XMLString); sb.Append("</root>");
            XElement xmlCounselType = XElement.Parse(sb.ToString());
            // 輔導方式--
            Dictionary<string, string> itemDict = new Dictionary<string, string>();
            if (xmlCounselType != null)
            {
                foreach (XElement elm in xmlCounselType.Elements("Item"))
                {
                    if (elm.Attribute("name") == null)
                        continue;

                    switch (elm.Attribute("name").Value.ToString())
                    {
                        // 暫時結案
                        case "暫時結案":
                            itemDict.Add("輔導方式:暫時結案","1");
                            break;

                        // 專案輔導
                        case "專案輔導":
                            itemDict.Add("輔導方式:專案輔導","1");
                            break;
                        // 導師輔導
                        case "導師輔導":
                            itemDict.Add("輔導方式:導師輔導","1");
                            break;
                        // 轉介**
                        case "轉介":
                            if (elm.Attribute("remark") != null)
                                itemDict.Add("輔導方式:轉介備註", elm.Attribute("remark").Value.ToString());
                            
                                itemDict.Add("輔導方式:轉介","1");
                            break;
                        // 就醫**
                        case "就醫":
                            if (elm.Attribute("remark") != null)
                                itemDict.Add("輔導方式:就醫備註" , elm.Attribute("remark").Value.ToString());
                            
                                itemDict.Add("輔導方式:就醫","1");
                            break;
                        // 其它**
                        case "其它":
                            if (elm.Attribute("remark") != null)
                                itemDict.Add("輔導方式:其它備註" , elm.Attribute("remark").Value.ToString());
                            
                                itemDict.Add("輔導方式:其它","1");
                            break;
                    }
                }
            }
            return itemDict;
        }

        /// <summary>
        /// 解析輔導歸類成Dict
        /// </summary>
        /// <param name="XMLString"></param>
        /// <returns></returns>
        public static Dictionary<string, string> GetConvertCounselXMLVal_CounselTypeKind(string XMLString)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<root>"); sb.Append(XMLString); sb.Append("</root>");
            XElement xmlCounselTypeKind = XElement.Parse(sb.ToString());
            Dictionary<string, string> itemDict = new Dictionary<string, string>();
            // 輔導歸類--
            if (xmlCounselTypeKind != null)
            {
                foreach (XElement elm in xmlCounselTypeKind.Elements("Item"))
                {
                    if (elm.Attribute("name") == null)
                        continue;

                    switch (elm.Attribute("name").Value.ToString())
                    {
                        // 違規
                        case "違規":
                            itemDict.Add("輔導歸類:違規","1");
                            break;

                        // 遲曠
                        case "遲曠":
                            itemDict.Add("輔導歸類:遲曠","1");
                            break;

                        // 學習
                        case "學習":
                            itemDict.Add("輔導歸類:學習","1");
                            break;

                        // 生涯
                        case "生涯":
                            itemDict.Add("輔導歸類:生涯","1");
                            break;

                        // 人
                        case "人":
                            itemDict.Add("輔導歸類:人","1");
                            break;

                        // 人際
                        case "人際":
                            itemDict.Add("輔導歸類:人際", "1");
                            break;


                        // 休退轉
                        case "休退轉":
                            itemDict.Add("輔導歸類:休退轉","1");
                            break;

                        // 家庭
                        case "家庭":
                            itemDict.Add("輔導歸類:家庭","1");
                            break;

                        // 師生
                        case "師生":
                            itemDict.Add("輔導歸類:師生","1");
                            break;

                        // 情感
                        case "情感":
                            itemDict.Add("輔導歸類:情感","1");
                            break;

                        // 精神
                        case "精神":
                            itemDict.Add("輔導歸類:精神","1");
                            break;

                        case "家暴": itemDict.Add("輔導歸類:家暴", "1"); break;
                        case "霸凌": itemDict.Add("輔導歸類:霸凌", "1"); break;
                        case "中輟": itemDict.Add("輔導歸類:中輟", "1"); break;
                        case "性議題": itemDict.Add("輔導歸類:性議題", "1"); break;
                        case "戒毒": itemDict.Add("輔導歸類:戒毒", "1"); break;
                        case "網路成癮": itemDict.Add("輔導歸類:網路成癮", "1"); break;
                        case "情緒障礙": itemDict.Add("輔導歸類:情緒障礙", "1"); break;

                        // 其它**
                        case "其它":
                            if (elm.Attribute("remark") != null)
                                itemDict.Add("輔導歸類:其它備註" , elm.Attribute("remark").Value.ToString());
                            
                                itemDict.Add("輔導歸類:其它","1");
                            break;

                    }
                }
            }
            return itemDict;
        }


        /// <summary>
        /// 取得所有學生ID,Key:(學號+姓名) Value:StudentID(int型態)
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, int> GetAllStudentID1()
        {
            Dictionary<string, int> retVal = new Dictionary<string, int>();
            QueryHelper hepler = new QueryHelper();
            // id,學號,姓名,狀態
            string strSQL = "select id,student_number,name,status from student;";
            DataTable dt = hepler.Select(strSQL);

            foreach (DataRow dr in dt.Rows)
            {
                int value = int.Parse(dr[0].ToString());
                string key = dr[1].ToString().Trim() + dr[2].ToString().Trim();
                if(!retVal.ContainsKey(key))
                    retVal.Add(key,value );
            }
            return retVal;
        }

        /// <summary>
        /// 取得所有學生輔導相關老師名稱(班導師與有設定輔導相關老師
        /// </summary>
        /// <returns></returns>
        public static Dictionary <string,Dictionary<string,int >> GetAllCounselStudentTeacherNameDict()
        {
            Dictionary<string,Dictionary<string,int>> retVal = new Dictionary<string,Dictionary<string,int>> ();

            // 班導師
            QueryHelper hepler = new QueryHelper();
            // 學生編號、教師名稱、教師暱稱、教師編號
            string strSQL = "select student.id,teacher.teacher_name,teacher.nickname,teacher.id as id2 from student inner join class on student.ref_class_id =class.id inner join teacher on class.ref_teacher_id = teacher.id;";
            DataTable dt = hepler.Select(strSQL);

            foreach (DataRow dr in dt.Rows)
            {
                // 學生編號
                string key = dr[0].ToString();

                string TeacherName = "";

                if (string.IsNullOrEmpty(dr[2].ToString()))
                    TeacherName = dr[1].ToString();
                else
                    TeacherName = dr[1].ToString() + dr[2].ToString();

                int TeacherID = int.Parse (dr[3].ToString());

                if (!retVal.ContainsKey(key))
                {
                    Dictionary<string, int> temp = new Dictionary<string, int>();
                    temp.Add(TeacherName, TeacherID);
                    retVal.Add(key,temp);
                }
            }

            // 有設定晤談相關教師
            QueryHelper hepler1 = new QueryHelper();
            // 學生編號、教師名稱、教師暱稱、教師編號
            string strSQL1 = "select counsel.student_list.ref_student_id,teacher.teacher_name,teacher.nickname,teacher.id as id2 from $counsel.student_list inner join tag_teacher on tag_teacher.id=$counsel.student_list.ref_teacher_tag_id inner join teacher on teacher.id=tag_teacher.ref_teacher_id";
            DataTable dt1 = hepler.Select(strSQL);

            foreach (DataRow dr in dt1.Rows)
            {
                // 學生編號
                string key = dr[0].ToString();

                string TeacherName = "";

                if (string.IsNullOrEmpty(dr[2].ToString()))
                    TeacherName = dr[1].ToString();
                else
                    TeacherName = dr[1].ToString() + dr[2].ToString();

                int TeacherID = int.Parse(dr[3].ToString());

                if (!retVal.ContainsKey(key))
                {
                    Dictionary<string, int> temp = new Dictionary<string, int>();
                    temp.Add(TeacherName, TeacherID);
                    retVal.Add(key, temp);
                }
                else
                    retVal[key].Add(TeacherName,TeacherID);
            }

            return retVal;
        }

        /// <summary>
        /// 取得輔導晤談勾選項目名稱
        /// </summary>
        /// <returns></returns>
        public static List<string> GetCounselXmlItemList_InterView()
        {
            List<string> retVal = new List<string>();
            retVal.Add("參與人員:學生");
            retVal.Add("參與人員:家長");
            retVal.Add("參與人員:專家");
            retVal.Add("參與人員:醫師");
            retVal.Add("參與人員:社工人員");
            retVal.Add("參與人員:導師");
            retVal.Add("參與人員:教官");
            retVal.Add("參與人員:輔導老師");
            retVal.Add("參與人員:任課老師");
            retVal.Add("參與人員:其它");
            retVal.Add("參與人員:其它備註");
            retVal.Add("輔導方式:暫時結案");
            retVal.Add("輔導方式:專案輔導");
            retVal.Add("輔導方式:導師輔導");
            retVal.Add("輔導方式:轉介");
            retVal.Add("輔導方式:轉介備註");
            retVal.Add("輔導方式:就醫");
            retVal.Add("輔導方式:就醫備註");
            retVal.Add("輔導方式:其它");
            retVal.Add("輔導方式:其它備註");
            retVal.Add("輔導歸類:違規");
            retVal.Add("輔導歸類:遲曠");
            retVal.Add("輔導歸類:學習");
            retVal.Add("輔導歸類:生涯");
            retVal.Add("輔導歸類:人際");
            retVal.Add("輔導歸類:休退轉");
            retVal.Add("輔導歸類:家庭");
            retVal.Add("輔導歸類:師生");
            retVal.Add("輔導歸類:情感");
            retVal.Add("輔導歸類:精神");
            retVal.Add("輔導歸類:家暴");
            retVal.Add("輔導歸類:霸凌");
            retVal.Add("輔導歸類:中輟");
            retVal.Add("輔導歸類:性議題");
            retVal.Add("輔導歸類:戒毒");
            retVal.Add("輔導歸類:網路成癮");
            retVal.Add("輔導歸類:情緒障礙");
            retVal.Add("輔導歸類:其它");
            retVal.Add("輔導歸類:其它備註");
            return retVal;
        }

        /// <summary>
        /// 個案會議選項
        /// </summary>
        /// <returns></returns>
        public static List<string> GetCounselXmlItemList_CaseMeeting()
        {
            List<string> retVal = new List<string>();
            retVal.Add("參與人員:學生");
            retVal.Add("參與人員:家長");
            retVal.Add("參與人員:專家");
            retVal.Add("參與人員:醫師");
            retVal.Add("參與人員:社工人員");
            retVal.Add("參與人員:導師");
            retVal.Add("參與人員:教官");
            retVal.Add("參與人員:輔導老師");
            retVal.Add("參與人員:任課老師");
            retVal.Add("參與人員:其它");
            retVal.Add("參與人員:其它備註");
            retVal.Add("輔導方式:暫時結案");
            retVal.Add("輔導方式:專案輔導");
            retVal.Add("輔導方式:導師輔導");
            retVal.Add("輔導方式:轉介");
            retVal.Add("輔導方式:轉介備註");
            retVal.Add("輔導方式:就醫");
            retVal.Add("輔導方式:就醫備註");
            retVal.Add("輔導方式:其它");
            retVal.Add("輔導方式:其它備註");
            retVal.Add("輔導歸類:違規");
            retVal.Add("輔導歸類:遲曠");
            retVal.Add("輔導歸類:學習");
            retVal.Add("輔導歸類:生涯");
            retVal.Add("輔導歸類:人");
            retVal.Add("輔導歸類:休退轉");
            retVal.Add("輔導歸類:家庭");
            retVal.Add("輔導歸類:師生");
            retVal.Add("輔導歸類:情感");
            retVal.Add("輔導歸類:精神");
            retVal.Add("輔導歸類:其它");
            retVal.Add("輔導歸類:其它備註");
            return retVal;
        }

        /// <summary>
        /// 依測驗編號取得已有測驗學生學號
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, string> GetHasCounselStudQuizDataStudNumberByQuizID(string QuizID)
        {
            Dictionary<string, string> retVal = new Dictionary<string, string>();
            QueryHelper qh = new QueryHelper();
            string strSQL = "select student.student_number,student.id from student inner join $counsel.student_quiz_data on  student.id=$counsel.student_quiz_data.ref_student_id where ref_quiz_uid="+QuizID;
            DataTable dt = qh.Select(strSQL);
            foreach (DataRow dr in dt.Rows)
            {
                string key = dr[0].ToString();
                if (!retVal.ContainsKey(key))
                    retVal.Add(key, dr[1].ToString());            
            }           

            return retVal;
        }

        /// <summary>
        /// 取得所有學生學號_狀態對應 StudentNumber_Status,id
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, int> GetAllStudenNumberStatusDict()
        {
            Dictionary<string, int> retVal = new Dictionary<string, int>();
            QueryHelper qh = new QueryHelper();
            string strSQL = "select student.student_number,student.status,student.id from student;";
            DataTable dt = qh.Select(strSQL);
            foreach (DataRow dr in dt.Rows)
            {
                string key = dr[0].ToString() + "_" + dr[1].ToString();
                int id = int.Parse(dr[2].ToString());
                if (!retVal.ContainsKey(key))
                    retVal.Add(key, id);
            }

            return retVal;        
        }

        /// <summary>
        /// 取得學生班級座號狀態對照 class_seatno_status,id，如有同狀態班座取第一筆
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, int> GetAllStudentClassSeatNoStatusDict()
        {
            Dictionary<string, int> retVal = new Dictionary<string, int>();
            QueryHelper qh = new QueryHelper();
            string strSQL = "select class.class_name,student.seat_no,student.status,student.id from student inner join class on student.ref_class_id=class.id where seat_no is not null order by student.status,class_name,seat_no";
            DataTable dt = qh.Select(strSQL);
            foreach (DataRow dr in dt.Rows)
            { 
                string key=dr[0].ToString()+"_"+dr[1].ToString()+"_"+dr[2].ToString();
                int id = int.Parse(dr[3].ToString());
                if (!retVal.ContainsKey(key))
                    retVal.Add(key, id);
            }

            return retVal;
        }

        /// <summary>
        /// 取得教師狀態為一般的名稱與ID對應 name(niclname),id
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, int> GetAllTeacherNameIDDict()
        {
            QueryHelper helper = new QueryHelper();

            // 非刪除教師
            string sqlStr1 = "select teacher.id,teacher.teacher_name,teacher.nickname from teacher where status=1;";
            DataTable dt1 = helper.Select(sqlStr1);
            Dictionary<string, int> retVal = new Dictionary<string, int>();
            foreach (DataRow dr in dt1.Rows)
            {
                int id = int.Parse(dr[0].ToString());

                string name = "";
                if (string.IsNullOrEmpty(dr[2].ToString()))
                    name = dr[1].ToString();
                else
                    name = dr[1].ToString() + "(" + dr[2].ToString() + ")";

                if (!retVal.ContainsKey(name))
                    retVal.Add(name, id);
            }
            return retVal;
        }

        /// <summary>
        /// 取得學生狀態與資料庫數字對應
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string,string> GetStudentStatusDBValDict()
        {
            Dictionary<string, string> retVal = new Dictionary<string, string>();
            retVal.Add("一般", "1");
            retVal.Add("延修", "2");
            retVal.Add("休學", "4");
            retVal.Add("輟學", "8");
            retVal.Add("畢業或離校", "16");
            retVal.Add("刪除", "256");
            retVal.Add("", "1");
            return retVal;
        }

        /// <summary>
        /// 透過學號狀態取得學生ID
        /// </summary>
        /// <param name="StudentNumber"></param>
        /// <param name="status"></param>
        /// <returns></returns>
        public static int GetStudentID(string StudentNumber,string status)
        {
            int retVal = 0;
            string key =StudentNumber + "_";
            if (Global._StudentStatusDBDict.ContainsKey(status))
                key += Global._StudentStatusDBDict[status];
            else
                key += "1";

            if (Global._AllStudentNumberStatusIDTemp.ContainsKey(key))
                retVal = Global._AllStudentNumberStatusIDTemp[key];

            return retVal;
        }

        /// <summary>
        /// 透過班級座號狀態取得學生ID
        /// </summary>
        /// <param name="StudentNumber"></param>
        /// <param name="status"></param>
        /// <returns></returns>
        public static int GetStudentID(string ClassName,string SeatNo, string status)
        {
            int retVal = 0;
            string key = ClassName+"_"+SeatNo + "_";
            if (Global._StudentStatusDBDict.ContainsKey(status))
                key += Global._StudentStatusDBDict[status];
            else
                key += "1";

            if (Global._AllStudentClassSeatNoDictTemp.ContainsKey(key))
                retVal = Global._AllStudentClassSeatNoDictTemp[key];

            return retVal;
        }

        /// <summary>
        /// 檢查是否加入資料項目
        /// </summary>
        /// <param name="code"></param>
        /// <returns></returns>
        public static bool CheckAddContent(string code)
        {
            bool retVal = false;
            if (FISCA.Permission.UserAcl.Current[code].Viewable)
                retVal = true;

            return retVal;
        }

        /// <summary>
        /// 取得系統內班級的年級
        /// </summary>
        /// <returns></returns>
        public static List<string> GetClassGradeYear()
        {
            List<string> retVal = new List<string>();
            QueryHelper qh = new QueryHelper();
            string strQuery = "select distinct grade_year from class where grade_year is not null order by grade_year";
            DataTable dt = qh.Select(strQuery);
            foreach (DataRow dr in dt.Rows)            
                retVal.Add(dr[0].ToString());    

            return retVal;        
        }

        /// <summary>
        /// 取得系統內班級的年級文字<string,int> 年級,數字
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string,int> GetClassGradeYearDict()
        {
            Dictionary<string, int> retVal = new Dictionary<string, int>();
            QueryHelper qh = new QueryHelper();
            string strQuery = "select distinct grade_year from class where grade_year is not null order by grade_year";
            DataTable dt = qh.Select(strQuery);
            foreach (DataRow dr in dt.Rows)
            {
                int gr;
                if (int.TryParse(dr[0].ToString(), out gr))
                {
                    string key = "";
                    if (gr > 12)
                        gr -= 12;

                    switch (gr)
                    {                        
                        case 1: key = "一年級"; break;
                        case 2: key = "二年級"; break;
                        case 3: key = "三年級"; break;
                        case 4: key = "四年級"; break;
                        case 5: key = "五年級"; break;
                        case 6: key = "六年級"; break;
                        case 7: key = "七年級"; break;
                        case 8: key = "八年級"; break;
                        case 9: key = "九年級"; break;
                        case 10: key = "十年級"; break;
                        case 11: key = "十一年級"; break;
                        case 12: key = "十二年級"; break;
                    }
                    if(!retVal.ContainsKey(key) && !String.IsNullOrEmpty(key))
                        retVal.Add(key, gr);
                }
            }
            return retVal;
        }

        /// <summary>
        /// 轉換年級學期字串(一上..)
        /// </summary>
        /// <param name="GradeYear"></param>
        /// <param name="Semester"></param>
        /// <returns></returns>
        public static string ConvertGradeYearSemester(int GradeYear, int Semester)
        {
            string retVal = "";

            switch (GradeYear)
            {
                case 1: retVal = "一"; break;
                case 2: retVal = "二"; break;
                case 3: retVal = "三"; break;
                case 4: retVal = "四"; break;
                case 5: retVal = "五"; break;
                case 6: retVal = "六"; break;
                case 7: retVal = "七"; break;
                case 8: retVal = "八"; break;
                case 9: retVal = "九"; break;
            }

            if (Semester == 1)
                retVal += "上";

            if (Semester == 2)
                retVal += "下";

            return retVal;
        }

        /// <summary>
        /// 讀取日常生活表現具體建議與導師評語
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string, List<DAO.ABCard_StudentTextScore>> GetABCard_StudentTextScoreDict(List<string> StudentIDList)
        {
            Dictionary<string, List<DAO.ABCard_StudentTextScore>> retVal = new Dictionary<string, List<DAO.ABCard_StudentTextScore>>();
            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string strSQL = "select ref_student_id,school_year,semester,sb_comment,text_score from sems_moral_score where ref_student_id in(" + string.Join(",", StudentIDList.ToArray()) + ")";
                DataTable dt = qh.Select(strSQL);

                // 讀取學期歷程
                List<SemesterHistoryRecord> SemesterHistoryRecordList = SemesterHistory.SelectByStudentIDs(StudentIDList);

                List<DAO.ABCard_StudentTextScore> dataList = new List<DAO.ABCard_StudentTextScore>();

                foreach(DataRow dr in dt.Rows)
                {
                    DAO.ABCard_StudentTextScore asts = new DAO.ABCard_StudentTextScore();
                    asts.StudentID = dr[0].ToString();
                    asts.SchoolYear = int.Parse(dr[1].ToString());
                    asts.Semester = int.Parse(dr[2].ToString());
                           
                    // 比對學期歷程填入年級
                    foreach (SemesterHistoryRecord shr in SemesterHistoryRecordList.Where(x => x.RefStudentID == asts.StudentID))
                        foreach (SemesterHistoryItem shi in (from data in shr.SemesterHistoryItems where data.SchoolYear == asts.SchoolYear && data.Semester == asts.Semester select data))
                            asts.GradeYear = shi.GradeYear;                    

                    asts.sb_Comment = dr[3].ToString();
                    // 解析資料
                    if (dr[4] != null && dr[4].ToString()!="")
                    {
                        XElement elm = XElement.Parse(dr[4].ToString());
                        if (elm != null)
                        {
                            if (elm.Element("DailyLifeRecommend") != null)
                                if(elm.Element("DailyLifeRecommend").Attribute("Description") !=null)
                                    asts.DailyLifeRecommend = elm.Element("DailyLifeRecommend").Attribute("Description").Value;
                        }
                    }
                    dataList.Add(asts);
                }

                // 填入並依年級學期排序
                foreach (string str in StudentIDList)                    
                    retVal.Add(str, (from data in dataList where data.StudentID== str orderby data.GradeYear,data.Semester select data).ToList());
            }
            return retVal;
        }

        /// <summary>
        /// 取得畢業資訊
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string, LeaveInfoRecord> GetStudentGraduateDict(List<string> StudentIDList)
        {
            List<LeaveInfoRecord> dataList = LeaveInfo.SelectByStudentIDs(StudentIDList);
            Dictionary<string, LeaveInfoRecord> retVal = new Dictionary<string, LeaveInfoRecord>();
            
            foreach (LeaveInfoRecord lir in dataList)
                retVal.Add(lir.RefStudentID, lir);

            return retVal;
        }

        /// <summary>
        /// 取得學生異動資料
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string, List<DAO.AB_StudDateString>> GetStudentUpdateRecDict(List<string> StudentIDList)
        {
            Dictionary<string, List<DAO.AB_StudDateString>> retVal = new Dictionary<string, List<DAO.AB_StudDateString>>();

            // 取得異動資料
            List<UpdateRecordRecord> updateRecList = UpdateRecord.SelectByStudentIDs(StudentIDList);

            List<DAO.AB_StudDateString> dataList = new List<DAO.AB_StudDateString>();
            // 轉換格式
            foreach (UpdateRecordRecord rec in updateRecList)
            {
                DAO.AB_StudDateString ass = new DAO.AB_StudDateString();
                ass.StudentID = rec.StudentID;
                ass.Date = DateTime.Parse(rec.UpdateDate);
                ass.Text1 = rec.UpdateCode;
                switch (rec.UpdateCode)
                {
                    case "1": ass.Text1 = "新生"; break;
                    case "2": ass.Text1 = "畢業"; break;
                    case "3": ass.Text1 = "轉入"; break;
                    case "4": ass.Text1 = "轉出"; break;
                    case "5": ass.Text1 = "休學"; break;
                    case "6": ass.Text1 = "復學"; break;
                    case "7": ass.Text1 = "中輟"; break;
                    case "8": ass.Text1 = "續讀"; break;
                    case "9": ass.Text1 = "更正學籍"; break;                
                }
                ass.Text2 = rec.UpdateDescription;
                dataList.Add(ass);
            }

            // 加入回傳值
            foreach (string str in StudentIDList)
                retVal.Add(str, (from data in dataList where data.StudentID == str orderby data.Date select data).ToList());

            return retVal;
        }

        /// <summary>
        /// 取得學生獎懲資料
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string, List<DAO.AB_StudDateString>> GetStudentMDRecDict(List<string> StudentIDList)
        {
            Dictionary<string, List<DAO.AB_StudDateString>> retVal = new Dictionary<string, List<DAO.AB_StudDateString>>();
            // 讀取獎勵
            List<MeritRecord> MeritList = Merit.SelectByStudentIDs(StudentIDList);

            // 讀取懲戒
            List<DemeritRecord> DemeritList = Demerit.SelectByStudentIDs(StudentIDList);

            List<DAO.AB_StudDateString> dataList = new List<DAO.AB_StudDateString>();

            List<string> tempString = new List<string>();
            // 解析獎勵資料            
            foreach (MeritRecord rec in MeritList)
            {
                tempString.Clear();
                DAO.AB_StudDateString ass = new DAO.AB_StudDateString();
                ass.StudentID = rec.RefStudentID;
                ass.Date = rec.OccurDate;
                if (rec.MeritA.HasValue && rec.MeritA.Value>0)
                    tempString.Add("大功:" + rec.MeritA.Value);

                if (rec.MeritB.HasValue && rec.MeritB.Value>0)
                    tempString.Add("小功:" + rec.MeritB.Value);

                if (rec.MeritC.HasValue && rec.MeritC.Value>0)
                    tempString.Add("獎勵:" + rec.MeritC.Value);

                if (rec.MeritFlag == "2")
                    tempString.Add("留校察看");

                if (tempString.Count > 0)
                    ass.Text1 = string.Join(",", tempString.ToArray());

                ass.Text2 = rec.Reason;

                dataList.Add(ass);
            }

            // 解析懲戒
            foreach (DemeritRecord rec in DemeritList)
            {
                tempString.Clear();
                DAO.AB_StudDateString ass = new DAO.AB_StudDateString();
                ass.StudentID = rec.RefStudentID;
                ass.Date = rec.OccurDate;

                if (rec.DemeritA.HasValue && rec.DemeritA.Value>0)
                    tempString.Add("大過:" + rec.DemeritA.Value);

                if (rec.DemeritB.HasValue && rec.DemeritB.Value>0)
                    tempString.Add("小過:" + rec.DemeritB.Value);

                if (rec.DemeritC.HasValue && rec.DemeritC.Value>0)
                    tempString.Add("警告:" + rec.DemeritC.Value);

                if (rec.ClearDate.HasValue)
                    tempString.Add("銷過日期:" + rec.ClearDate.Value + ",銷過事由:" + rec.ClearReason);

                if (tempString.Count > 0)
                    ass.Text1 = string.Join(",", tempString.ToArray());

                ass.Text2 = rec.Reason;

                dataList.Add(ass);
            
            }

            // 放入回傳資料並依日期排序
            foreach (string str in StudentIDList)
                retVal.Add(str, (from data in dataList where data.StudentID == str orderby data.Date select data).ToList());
            
            return retVal;        
        }

        /// <summary>
        /// 取得AB表用學生選社團
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string, List<DAO.AB_StudText>> GetABCard_StudentSpecCourseDict(List<string> StudentIDList)
        {
            Dictionary<string, List<DAO.AB_StudText>> retVal = new Dictionary<string, List<DAO.AB_StudText>>();
            if (StudentIDList.Count > 0)
            {
                // 取得學生選社
                QueryHelper qh = new QueryHelper();
                string strSQL = "select ref_student_id,course.school_year,course.semester,course.course_name from sc_attend inner join course on sc_attend.ref_course_id=course.id inner join tag_course on tag_course.ref_course_id=sc_attend.ref_course_id and tag_course.ref_tag_id in(select id from Tag where name in('社團','聯課活動') and prefix in('社團','聯課活動')) where ref_student_id in(" + string.Join(",", StudentIDList.ToArray()) + ")";
                DataTable dt = qh.Select(strSQL);

                // 讀取學期歷程
                List<SemesterHistoryRecord> SemesterHistoryRecordList = SemesterHistory.SelectByStudentIDs(StudentIDList);
                List<DAO.AB_StudText> dataList = new List<DAO.AB_StudText>();
                foreach (DataRow dr in dt.Rows)
                {
                    DAO.AB_StudText ast = new DAO.AB_StudText();
                    ast.StudentID = dr[0].ToString();
                    ast.SchoolYear = int.Parse(dr[1].ToString());
                    ast.Semester = int.Parse(dr[2].ToString());
                    ast.Text1 = dr[3].ToString();

                    // 比對學期歷程填入年級
                    foreach (SemesterHistoryRecord shr in SemesterHistoryRecordList.Where(x => x.RefStudentID == ast.StudentID))
                        foreach (SemesterHistoryItem shi in (from data in shr.SemesterHistoryItems where data.SchoolYear == ast.SchoolYear && data.Semester == ast.Semester select data))
                            ast.GradeYear = shi.GradeYear;

                    dataList.Add(ast);
                }

                foreach(string str in StudentIDList)
                    retVal.Add(str,(from data in dataList where data.StudentID== str orderby data.GradeYear,data.Semester select data).ToList());
            }
            return retVal;
        
        }

        /// <summary>
        /// 取得AB表用學生擔任幹部
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string, List<DAO.AB_StudText>> GetABCard_StudentTheCadreDict(List<string> StudentIDList)
        {
            Dictionary<string, List<DAO.AB_StudText>> retVal = new Dictionary<string, List<DAO.AB_StudText>>();
            if (StudentIDList.Count > 0)
            {
                // 取得幹部資料
                QueryHelper qh = new QueryHelper();
                string strSQL = "select studentid,schoolyear,semester,referencetype,cadrename from $behavior.thecadre where studentid in('" + string.Join("','", StudentIDList.ToArray()) + "')";
                DataTable dt = qh.Select(strSQL);

                // 讀取學期歷程
                List<SemesterHistoryRecord> SemesterHistoryRecordList = SemesterHistory.SelectByStudentIDs(StudentIDList);

                List<DAO.AB_StudText> dataList = new List<DAO.AB_StudText>();
                foreach (DataRow dr in dt.Rows)
                {
                    DAO.AB_StudText ast = new DAO.AB_StudText();
                    ast.StudentID = dr[0].ToString();
                    ast.SchoolYear = int.Parse(dr[1].ToString());
                    ast.Semester = int.Parse(dr[2].ToString());
                    ast.Text1 = dr[3].ToString() + ":" + dr[4].ToString();

                    // 比對學期歷程填入年級
                    foreach (SemesterHistoryRecord shr in SemesterHistoryRecordList.Where(x => x.RefStudentID == ast.StudentID))
                        foreach (SemesterHistoryItem shi in (from data in shr.SemesterHistoryItems where data.SchoolYear == ast.SchoolYear && data.Semester == ast.Semester select data))
                            ast.GradeYear = shi.GradeYear;

                    dataList.Add(ast);
                }

                foreach (string str in StudentIDList)
                    retVal.Add(str, (from data in dataList where data.StudentID == str orderby data.GradeYear, data.Semester select data).ToList());
            }
            return retVal;

        }

        /// <summary>
        /// AB表用取得學生學期領域成績與畢業成績
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string,DAO.AB_StudSemsDomainScore> GetAB_StudSemsDomainScoreDict(List<string> StudentIDList)
        {           

            Dictionary<string, DAO.AB_StudSemsDomainScore> retVal = new Dictionary<string, DAO.AB_StudSemsDomainScore>();

            if (StudentIDList.Count > 0)
            {
                // 讀取學期歷程
                List<SemesterHistoryRecord> SemesterHistoryRecordList = SemesterHistory.SelectByStudentIDs(StudentIDList);

                // 讀取國中領域成績
                List<JHSemesterScoreRecord> SemesterScoreRecordList = JHSemesterScore.SelectByStudentIDs(StudentIDList);

               // 取得畢業成績
                Dictionary<string, Dictionary<string, decimal?>> GraduateScoreDict = GetStudGraduateDictJH(StudentIDList);
                
               // 轉換資料
                foreach (string studID in StudentIDList)
                {
                    DAO.AB_StudSemsDomainScore abssds = new DAO.AB_StudSemsDomainScore();
                    abssds.StudentID = studID;
                    List<DAO.AB_RptColIdx> GradeYearList = new List<DAO.AB_RptColIdx>();
                    // 放入學年度學期年級對照，儲存時依照年級學期排序
                    foreach (SemesterHistoryRecord shr in SemesterHistoryRecordList.Where(x => x.RefStudentID == studID))
                        foreach (SemesterHistoryItem shi in shr.SemesterHistoryItems)
                        {
                            DAO.AB_RptColIdx abrci = new DAO.AB_RptColIdx();
                            abrci.SchoolYear = shi.SchoolYear;
                            abrci.Semester = shi.Semester;
                            abrci.GradeYear = shi.GradeYear;
                            GradeYearList.Add(abrci);
                        }
                    abssds.GradeYearList = (from data in GradeYearList orderby data.GradeYear, data.Semester select data).ToList();


                    // 建立學期領域成績對照
                    // 取得這位學生學期成績                    
                    foreach (JHSemesterScoreRecord scoreRec in SemesterScoreRecordList.Where(x => x.RefStudentID == studID))
                    {
                        // 當比對到學年度學期相同才放入
                        foreach (DAO.AB_RptColIdx abrci in (from data in abssds.GradeYearList where data.SchoolYear==scoreRec.SchoolYear && data.Semester == scoreRec.Semester select data))
                        {
                            // 是否有語文領域成績
                            bool hasLangScore = false;
                            
                            // 讀取領域成績                            
                            foreach(DomainScore score in scoreRec.Domains.Values)
                            {
                                // 過濾彈性課程領域

                                if (score.Domain == "彈性課程")
                                    continue;

                                // 加入領域名稱
                                if(!abssds.DomainNameDict.ContainsKey(score.Domain))
                                    abssds.DomainNameDict.Add(score.Domain, new List<string>());

                                DAO.AB_SemsDomainScore tmpSemsScore = new DAO.AB_SemsDomainScore();
                                if (score.Domain == "語文")
                                {
                                    hasLangScore = true;
                                    tmpSemsScore.GroupName = "語文";
                                }
                                tmpSemsScore.DomainName = score.Domain;                                
                                tmpSemsScore.DomainScore = score.Score;
                                tmpSemsScore.GradeYear = abrci.GradeYear;
                                tmpSemsScore.SchoolYear = score.SchoolYear;
                                tmpSemsScore.Semester = score.Semester;
                                abssds.SemsDomainScoreList.Add(tmpSemsScore);
                            }

                            // 處理新竹語文對應問題
                            // 讀取科目語文領域成績
                            //沒有語文成績，當科目有屬於語文成績，放入當作語文成績需要展開處理
                            if (hasLangScore == false)
                            {
                                Dictionary<string, decimal?> subjScoreDict = new Dictionary<string,decimal?> ();

                                foreach (SubjectScore Sscore in scoreRec.Subjects.Values.Where(x => x.Domain == "語文"))
                                    subjScoreDict.Add(Sscore.Subject, Sscore.Score);
                                
                                // 有科目有語文成績
                                if (subjScoreDict.Count > 0)
                                {
                                    DAO.AB_SemsDomainScore tmpSemsSocre = new DAO.AB_SemsDomainScore();
                                    tmpSemsSocre.DomainName = "語文";
                                    tmpSemsSocre.DomainScore = null;
                                    tmpSemsSocre.GradeYear = abrci.GradeYear;
                                    tmpSemsSocre.GroupName = "語文";
                                    tmpSemsSocre.LangSubjDict = subjScoreDict;
                                    tmpSemsSocre.SchoolYear = abrci.SchoolYear;
                                    tmpSemsSocre.Semester = abrci.Semester;
                                    abssds.SemsDomainScoreList.Add(tmpSemsSocre);
                                    if(!abssds.DomainNameDict.ContainsKey("語文"))
                                        abssds.DomainNameDict.Add("語文", subjScoreDict.Keys.ToList());
                                }
                            }    // 判斷語文                        

                            // 取得畢業成績
                            if (GraduateScoreDict.ContainsKey(studID))
                                abssds.GraduateScoreDict = GraduateScoreDict[studID];


                        } // 當比對到學年度學期相同才放入
                    } // 取得這位學生學期成績
                    retVal.Add(abssds.StudentID,abssds);
                }
            }
            return retVal;
        }

        /// <summary>
        /// 取得學生畢業領域成績(國中)
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string, Dictionary<string, decimal?>> GetStudGraduateDictJH(List<string> StudentIDList)
        {
            Dictionary<string, Dictionary<string, decimal?>> retVal = new Dictionary<string, Dictionary<string, decimal?>>();
            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string strSQL = "select id,grad_score from student where id in('" + string.Join("','", StudentIDList.ToArray()) + "')";
                DataTable dt = qh.Select(strSQL);

                foreach (DataRow dr in dt.Rows)
                {
                    string id = dr[0].ToString();
                    string strXml = dr[1].ToString();
                    Dictionary<string, decimal?> domainScoreDict = new Dictionary<string, decimal?>();
                    if (strXml != "")
                    {
                        XElement elms = XElement.Parse(strXml);                        
                        foreach (XElement elm in elms.Elements("Domain"))
                        {
                            decimal? score=null;
                            decimal s;
                            string name = elm.Attribute("Name").Value;
                            if (elm.Attribute("Score") != null)                            
                                if (decimal.TryParse(elm.Attribute("Score").Value, out s))
                                    score = s;
                            
                            
                            if (!domainScoreDict.ContainsKey(name))
                                domainScoreDict.Add(name, score);
                        }
                    }
                    retVal.Add(id, domainScoreDict);                
                }

            }
            return retVal;
        }

        /// <summary>
        /// 處理 Aspose Word 水平合併(左至右)(單一)
        /// </summary>
        /// <param name="left"></param>
        /// <param name="right"></param>
        public static void HorizontallyMergeCells(Cell left, Cell right)
        {
            left.CellFormat.HorizontalMerge = CellMerge.First;

            foreach (Node child in right.ChildNodes)
                left.AppendChild(child);

            right.CellFormat.HorizontalMerge = CellMerge.Previous;
            
            // 處理移除合併後有 \r\n 問題
            if (left.HasChildNodes)
                left.RemoveChild(left.LastChild);        
        }

        /// <summary>
        /// Aspose Word 垂直合併上至下(單一)
        /// </summary>
        /// <param name="top"></param>
        /// <param name="bottom"></param>
        public static void VerticallyMergeCells(Cell top, Cell bottom)
        {
            top.CellFormat.VerticalMerge = CellMerge.First;

            foreach (Node child in bottom.ChildNodes)
                top.AppendChild(child);

            bottom.CellFormat.VerticalMerge = CellMerge.Previous;

            if (top.HasChildNodes)
                top.RemoveChild(top.LastChild);
        }

        /// <summary>
        /// 轉換個案會議資料至AB表輔導資料
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string,List<DAO.AB_RptCounselData>> ParseCaseMeetingRecordToABRptData(List<string> StudentIDList)
        { 
            Dictionary<string,List<DAO.AB_RptCounselData>> retVal = new Dictionary<string,List<DAO.AB_RptCounselData>> ();
            if (StudentIDList.Count > 0)
            {
                DAO.UDTTransfer _UDTTransfer = new DAO.UDTTransfer();
                List<DAO.UDT_CounselCaseMeetingRecordDef> CounselCaseMeetingRecordList = _UDTTransfer.GetCaseMeetingRecordListByStudentIDList(StudentIDList);
                // 取得教師ID與名稱對照
                Dictionary<int, string> TeacherNameDict = Utility.GetTeacherIDNameDict();
                // 參與人員
                Dictionary<string, string> item_AttendessDict = new Dictionary<string, string>();
                // 輔導方式
                Dictionary<string, string> item_CounselTypeDict = new Dictionary<string, string>();
                // 輔導歸類
                Dictionary<string, string> item_CounselTypeKindDict = new Dictionary<string, string>();

                // 解析並排序
                foreach (string studId in StudentIDList)
                {
                    int sid=int.Parse(studId);
                    List<DAO.AB_RptCounselData> studData = new List<DAO.AB_RptCounselData> ();
                    foreach (DAO.UDT_CounselCaseMeetingRecordDef ccmrd in (from data in CounselCaseMeetingRecordList where data.StudentID == sid orderby data.MeetingDate select data ))
                    {
                        DAO.AB_RptCounselData data = new DAO.AB_RptCounselData ();
                        // 取得 XML 解析後
                        item_AttendessDict = GetConvertCounselXMLVal_Attendees(ccmrd.Attendees);
                        item_CounselTypeDict = GetConvertCounselXMLVal_CounselType(ccmrd.CounselType);
                        item_CounselTypeKindDict = GetConvertCounselXMLVal_CounselTypeKind(ccmrd.CounselTypeKind);

                        data.StudentID = studId;
                        
                        if(ccmrd.MeetingDate.HasValue)
                            data.DataDict.Add("會議日期", ccmrd.MeetingDate.Value.ToShortDateString());

                        data.DataDict.Add("會議事由", ccmrd.MeetingCause);
                        data.DataDict.Add("記錄者帳號", ccmrd.AuthorID);                      

                        // 記錄其它
                        data.OtherDataList = AB_Counset_AddString("個案編號", ccmrd.CaseNo, data.OtherDataList);
                        data.OtherDataList = AB_Counset_AddString("內容要點", ccmrd.ContentDigest, data.OtherDataList);
                        data.OtherDataList = AB_Counset_AddString("會議時間", ccmrd.MeetigTime, data.OtherDataList);
                        data.OtherDataList = AB_Counset_AddString("晤談地點", ccmrd.Place, data.OtherDataList);
                        if (TeacherNameDict.ContainsKey(ccmrd.CounselTeacherID))
                            data.OtherDataList=AB_Counset_AddString("晤談老師", TeacherNameDict[ccmrd.CounselTeacherID],data.OtherDataList);
                        
                        data.OtherDataList.AddRange(AB_Counsel_Item_Add(item_AttendessDict));

                        data.OtherDataList.AddRange(AB_Counsel_Item_Add(item_CounselTypeDict));

                        data.OtherDataList.AddRange(AB_Counsel_Item_Add(item_CounselTypeKindDict));
                        
                        studData.Add(data);                       
                    }
                    retVal.Add(studId,studData);
                }
            }
            return retVal;        
        }

        /// <summary>
        /// 解析AB表內輔導需要項目內重並放入
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public static List<string> AB_Counsel_Item_Add(Dictionary<string,string> source)
        {
            List<string> retVal = new List<string>();

            foreach (KeyValuePair<string, string> str in source)
            {
                if (!string.IsNullOrEmpty(str.Value))
                {
                    if (str.Value == "1")
                        retVal.Add(str.Key);
                    else  // 當有其它
                        retVal.Add(str.Key + ":" + str.Value);
                }            
            }
            return retVal;        
        }

        /// <summary>
        /// AB表報表用當有值才加入內容
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        /// <param name="source"></param>
        /// <returns></returns>
        public static List<string> AB_Counset_AddString(string name, string value,List<string> source)
        {
            if (!string.IsNullOrEmpty(value))
                source.Add(name + ":" + value);

            return source;
        }


        /// <summary>
        /// 轉換優先關懷資料至AB表輔導資料
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string, List<DAO.AB_RptCounselData>> ParseCareRecordToABRptData(List<string> StudentIDList)
        {
            Dictionary<string, List<DAO.AB_RptCounselData>> retVal = new Dictionary<string, List<DAO.AB_RptCounselData>>();
            if (StudentIDList.Count > 0)
            {
                DAO.UDTTransfer _UDTTransfer = new DAO.UDTTransfer();
                List<DAO.UDT_CounselCareRecordDef> _CounselCareRecordList = _UDTTransfer.GetCareRecordsByStudentIDList(StudentIDList);
                                
                foreach (string studId in StudentIDList)
                {
                    int sid = int.Parse(studId);
                    List<DAO.AB_RptCounselData> studData = new List<DAO.AB_RptCounselData>();
                    foreach (DAO.UDT_CounselCareRecordDef ccrd in (from data in _CounselCareRecordList where data.StudentID== sid orderby data.FileDate select data))
                    {
                        DAO.AB_RptCounselData data = new DAO.AB_RptCounselData();
                        if (ccrd.FileDate.HasValue)
                            data.DataDict.Add("立案日期", ccrd.FileDate.Value.ToShortDateString());
                        
                        data.StudentID = studId;
                        
                        data.DataDict.Add("個案類別", ccrd.CaseCategory);
                        data.DataDict.Add("記錄者帳號",ccrd.AuthorID);

                        data.OtherDataList = AB_Counset_AddString("代號", ccrd.CodeName, data.OtherDataList);
                        data.OtherDataList = AB_Counset_AddString("個案類別備註", ccrd.CaseCategoryRemark, data.OtherDataList);
                        data.OtherDataList = AB_Counset_AddString("個案來源", ccrd.CaseOrigin, data.OtherDataList);
                        data.OtherDataList = AB_Counset_AddString("個案來源備註", ccrd.CaseOriginRemark, data.OtherDataList);
                        data.OtherDataList = AB_Counset_AddString("優勢能力及財力", ccrd.Superiority, data.OtherDataList);
                        data.OtherDataList = AB_Counset_AddString("弱勢能力及財力", ccrd.Weakness, data.OtherDataList);
                        data.OtherDataList = AB_Counset_AddString("輔導人員輔導目標", ccrd.CounselGoal, data.OtherDataList);
                        data.OtherDataList = AB_Counset_AddString("校外協輔機構", ccrd.OtherInstitute, data.OtherDataList);
                        data.OtherDataList = AB_Counset_AddString("輔導人員輔導方式", ccrd.CounselType, data.OtherDataList);
                        data.OtherDataList = AB_Counset_AddString("協同輔導人員協助導師事項", ccrd.AssistedMatter, data.OtherDataList);

                        studData.Add(data);
                    }

                    retVal.Add(studId, studData);
                }
            }
            return retVal;        
        }

        /// <summary>
        /// 轉換晤談資料至AB表輔導資料
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string, List<DAO.AB_RptCounselData>> ParseInterViewDataToABRptData(List<string> StudentIDList)
        {
            Dictionary<string, List<DAO.AB_RptCounselData>> retVal = new Dictionary<string, List<DAO.AB_RptCounselData>>();
            if (StudentIDList.Count > 0)
            {
                DAO.UDTTransfer _UDTTransfer = new DAO.UDTTransfer();
                // 依學生ID取得晤談紀錄
                List<DAO.UDT_CounselStudentInterviewRecordDef> _CounselCareRecordList = _UDTTransfer.GetCounselStudentInterviewRecordByStudentIDList(StudentIDList);
                 // 取得教師ID與名稱對照
                Dictionary<int, string> TeacherNameDict = Utility.GetTeacherIDNameDict();

                Dictionary<string, string> item_AttendessDict = new Dictionary<string, string>();
                Dictionary<string, string> item_CounselTypeDict = new Dictionary<string, string>();
                Dictionary<string, string> item_CounselTypeKindDict = new Dictionary<string, string>();

                foreach (string studId in StudentIDList)
                {
                    int sid = int.Parse(studId);
                    List<DAO.AB_RptCounselData> studData = new List<DAO.AB_RptCounselData>();
                    foreach (DAO.UDT_CounselStudentInterviewRecordDef csir in (from data in _CounselCareRecordList where data.StudentID==sid orderby data.InterviewDate select data))
                    {
                        DAO.AB_RptCounselData data = new DAO.AB_RptCounselData();
                        // 取得 XML 解析後
                        item_AttendessDict = GetConvertCounselXMLVal_Attendees(csir.Attendees);
                        item_CounselTypeDict = GetConvertCounselXMLVal_CounselType(csir.CounselType);
                        item_CounselTypeKindDict = GetConvertCounselXMLVal_CounselTypeKind(csir.CounselTypeKind);

                        if(csir.InterviewDate.HasValue)
                            data.DataDict.Add("晤談日期",csir.InterviewDate.Value.ToShortDateString());

                        data.DataDict.Add("晤談對象", csir.IntervieweeType);
                        data.DataDict.Add("晤談事由", csir.Cause);
                        data.DataDict.Add("記錄者帳號", csir.AuthorID);

                        data.OtherDataList = AB_Counset_AddString("晤談編號", csir.InterviewNo, data.OtherDataList);
                        
                        if(TeacherNameDict.ContainsKey(csir.TeacherID))
                            data.OtherDataList = AB_Counset_AddString("晤談老師", TeacherNameDict[csir.TeacherID], data.OtherDataList);
                        
                        data.OtherDataList = AB_Counset_AddString("晤談方式", csir.InterviewType, data.OtherDataList);
                        data.OtherDataList = AB_Counset_AddString("晤談時間", csir.InterviewTime, data.OtherDataList);
                        data.OtherDataList = AB_Counset_AddString("晤談地點", csir.Place, data.OtherDataList);
                        data.OtherDataList = AB_Counset_AddString("內容要點", csir.ContentDigest, data.OtherDataList);

                        data.OtherDataList.AddRange(AB_Counsel_Item_Add(item_AttendessDict));

                        data.OtherDataList.AddRange(AB_Counsel_Item_Add(item_CounselTypeDict));

                        data.OtherDataList.AddRange(AB_Counsel_Item_Add(item_CounselTypeKindDict));

                        studData.Add(data);
                    }
                    retVal.Add(studId, studData);
                }
            }
            return retVal;        
        }

        /// <summary>
        /// 取得一般狀態的教師名稱帳號
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, string> GetTeacherNameLoginIDStatus1()
        {
            Dictionary<string, string> retVal = new Dictionary<string, string>();
            QueryHelper qh = new QueryHelper();            
            string strSQL = "select teacher_name||nickname as teachername,st_login_name from teacher where st_login_name is not null and status=1";
            DataTable dt = qh.Select(strSQL);

            foreach (DataRow dr in dt.Rows)
            {
                string key = dr[0].ToString();

                if (!retVal.ContainsKey(key))
                    retVal.Add(key, dr[1].ToString());            
            }

            return retVal;
        }

        /// <summary>
        /// 轉換Log需要班級座號姓名
        /// </summary>
        /// <param name="rec"></param>
        /// <returns></returns>
        public static string ConvertString1(StudentRecord rec)
        {
            List<string> retList = new List<string>();

            retList.Add("學號：" + rec.StudentNumber);
            if(rec.Class!=null)
                retList.Add("班級：" + rec.Class.Name);

            if (rec.SeatNo.HasValue)
                retList.Add("座號：" + rec.SeatNo.Value);

            retList.Add("姓名：" + rec.Name);

            if (retList.Count > 0)
                return string.Join(",", retList.ToArray());
            else
                return "";
        }

        /// <summary>
        /// 取得 Log 用名稱
        /// </summary>
        /// <param name="intStudentIDList"></param>
        /// <returns></returns>
        public static  Dictionary<int, string> GetConvertStringDict1fromDB(List<int> intStudentIDList)        
        {            
            Dictionary<int, string> retVal = new Dictionary<int, string>();
            QueryHelper qh = new QueryHelper();
            if (intStudentIDList.Count > 0)
            {
                string strSQL = "select student.id,'學號：'||student_number|| case when class.id is null then '' else ' , 班級：'||class.class_name end || case when seat_no is null then '' else ' , 座號：'||seat_no end||' , 姓名：'||student.name as name from student left join class on student.ref_class_id=class.id where student.id in(" + string.Join(",", intStudentIDList.ToArray()) + ")";
                DataTable dt = qh.Select(strSQL);
                foreach (DataRow dr in dt.Rows)
                {
                    int id = int.Parse(dr[0].ToString());
                    retVal.Add(id, dr[1].ToString());
                }
            }
            return retVal;
        }

        /// <summary>
        /// 將學生 ID 依照班級順序、班級名稱、座號排序
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static List<string> SortStudentID1(List<string> StudentIDList)
        {
            List<string> retVal = new List<string>();
            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string strSQL = "select student.id from student left join class on student.ref_class_id=class.id where student.id in("+string.Join(",",StudentIDList.ToArray())+") order by class.display_order asc ,class_name asc,student.seat_no asc;";
                DataTable dt = qh.Select(strSQL);

                foreach (DataRow dr in dt.Rows)
                    retVal.Add(dr[0].ToString());
            }
            return retVal;
        }

        /// <summary>
        /// 將學生 ID 依照班級順序、班級名稱、座號排序
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static List<int> SortStudentID2(List<int> StudentIDList)
        {
            List<int> retVal = new List<int>();
            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string strSQL = "select student.id from student left join class on student.ref_class_id=class.id where student.id in(" + string.Join(",", StudentIDList.ToArray()) + ") order by class.display_order asc ,class_name asc,student.seat_no asc;";
                DataTable dt = qh.Select(strSQL);

                foreach (DataRow dr in dt.Rows)
                    retVal.Add(int.Parse(dr[0].ToString()));
            }
            return retVal;
        }

        /// <summary>
        /// 透過班級ID取得班級內學生狀態一般的學生ID
        /// </summary>
        /// <param name="ClassIDList"></param>
        /// <returns></returns>
        public static List<string> GetStudentIDList1ByClassID(List<string> ClassIDList)
        {
            List<string> retVal = new List<string>();
            if (ClassIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string strSQL = "select student.id from student inner join class on student.ref_class_id=class.id where student.status in(1) and class.id in(" + string.Join(",", ClassIDList.ToArray()) + ") order by student_number;";
                DataTable dt = qh.Select(strSQL);
                foreach (DataRow dr in dt.Rows)
                    retVal.Add(dr[0].ToString());
            }
            return retVal;
        }

        /// <summary>
        /// 班座排序
        /// </summary>
        /// <param name="StudIDList"></param>
        /// <returns></returns>
        public static List<string> GetStudentIDListByStudentID(List<string> StudIDList)
        {
            List<string> retVal = new List<string>();
            if (StudIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string strSQL = "select student.id from student left join class on student.ref_class_id=class.id where student.id in(" + string.Join(",", StudIDList.ToArray()) + ") order by class.class_name,student.seat_no;";
                DataTable dt = qh.Select(strSQL);
                foreach (DataRow dr in dt.Rows)
                    retVal.Add(dr[0].ToString());
            }
            return retVal;
        }

        /// <summary>
        /// 報表用解析輔導綜合紀錄學年型 一：；二：；..
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public static string Counsel_Yearly_Parse1(UDTYearlyDataDef data)
        {
            List<string> retVal = new List<string>();
            if (data != null)
            {
                if (!string.IsNullOrWhiteSpace(data.G1))
                    retVal.Add("一："+data.G1);

                if (!string.IsNullOrWhiteSpace(data.G2))
                    retVal.Add("二："+data.G2);

                if (!string.IsNullOrWhiteSpace(data.G3))
                    retVal.Add("三："+data.G3);

                if (!string.IsNullOrWhiteSpace(data.G4))
                    retVal.Add("四："+data.G4);

                if (!string.IsNullOrWhiteSpace(data.G5))
                    retVal.Add("五："+data.G5);

                if (!string.IsNullOrWhiteSpace(data.G6))
                    retVal.Add("六："+data.G6);
            }
            return string.Join("；", retVal.ToArray());
        }

        /// <summary>
        /// 報表用解析輔導綜合紀錄學年型
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public static string Counsel_Yearly_Parse2(UDTYearlyDataDef data,int gr)
        {
            string retVal = "";
            if (data != null)
            {
                if (gr==1 && !string.IsNullOrWhiteSpace(data.G1))
                    retVal= data.G1;

                if (gr==2 && !string.IsNullOrWhiteSpace(data.G2))
                    retVal= data.G2;

                if (gr==3 && !string.IsNullOrWhiteSpace(data.G3))
                    retVal= data.G3;

                if (gr==4 && !string.IsNullOrWhiteSpace(data.G4))
                    retVal= data.G4;

                if (gr==5 && !string.IsNullOrWhiteSpace(data.G5))
                    retVal= data.G5;

                if (gr==6 && !string.IsNullOrWhiteSpace(data.G6))
                    retVal= data.G6;
            }
            return retVal;
        }
        /// <summary>
        /// 輔導綜合紀錄表報表多選轉換，點隔開
        /// </summary>
        /// <param name="data"></param>
        /// <param name="key"></param>
        /// <returns></returns>
        public static string Counsel_Multiple_Parse(List<UDTMultipleRecordDef> dataList,string key)
        {
            List<string> retVal = new List<string>();
            foreach (UDTMultipleRecordDef data in dataList.Where(x => x.Key == key))
            {
                if (string.IsNullOrWhiteSpace(data.Remark))
                    retVal.Add(data.Data);
                else
                    retVal.Add(data.Data + "：" + data.Remark);
            }

            return string.Join("，", retVal.ToArray());
        }

        /// <summary>
        /// 報表用解析輔導綜合紀錄優先順序 一：；二：；..
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public static string Counsel_PriorityData_Parse1(UDTPriorityDataDef data)
        {
            List<string> retVal = new List<string>();
            if (data != null)
            {
                if (!string.IsNullOrWhiteSpace(data.P1))
                    retVal.Add("一：" + data.P1);

                if (!string.IsNullOrWhiteSpace(data.P2))
                    retVal.Add("二：" + data.P2);

                if (!string.IsNullOrWhiteSpace(data.P3))
                    retVal.Add("三：" + data.P3);

                if (!string.IsNullOrWhiteSpace(data.P4))
                    retVal.Add("四：" + data.P4);

                if (!string.IsNullOrWhiteSpace(data.P5))
                    retVal.Add("五：" + data.P5);

                if (!string.IsNullOrWhiteSpace(data.P6))
                    retVal.Add("六：" + data.P6);
            }
            return string.Join("；", retVal.ToArray());
        }

        /// <summary>
        /// 處理報表學期型值
        /// </summary>
        /// <param name="data"></param>
        /// <param name="idx"></param>
        /// <returns></returns>
        public static string Counsel_SemesterData_Parse1(UDTSemesterDataDef data, int idx)
        {
            string retVal = "";
            if (data != null)
            {
                if (idx == 1 && !string.IsNullOrWhiteSpace(data.S1a))
                    retVal = data.S1a;

                if (idx == 2 && !string.IsNullOrWhiteSpace(data.S1b))
                    retVal = data.S1b;

                if (idx == 3 && !string.IsNullOrWhiteSpace(data.S2a))
                    retVal = data.S2a;

                if (idx == 4 && !string.IsNullOrWhiteSpace(data.S2b))
                    retVal = data.S2b;

                if (idx == 5 && !string.IsNullOrWhiteSpace(data.S3a))
                    retVal = data.S3a;

                if (idx == 6 && !string.IsNullOrWhiteSpace(data.S3b))
                    retVal = data.S3b;

                if (idx == 7 && !string.IsNullOrWhiteSpace(data.S4b))
                    retVal = data.S4b;

                if (idx == 8 && !string.IsNullOrWhiteSpace(data.S4b))
                    retVal = data.S4b;

                if (idx == 9 && !string.IsNullOrWhiteSpace(data.S5b))
                    retVal = data.S5b;

                if (idx == 10 && !string.IsNullOrWhiteSpace(data.S5b))
                    retVal = data.S5b;

                if (idx == 11 && !string.IsNullOrWhiteSpace(data.S6b))
                    retVal = data.S6b;

                if (idx == 12 && !string.IsNullOrWhiteSpace(data.S6b))
                    retVal = data.S6b;
            }
            return retVal;
        }

        /// <summary>
        /// 解析心理測驗
        /// </summary>
        /// <param name="strXML"></param>
        /// <returns></returns>
        public static string CounselStudQuizXmlParse1(string strXML)
        {
            List<string> retVal = new List<string>();
            if (!string.IsNullOrWhiteSpace(strXML))
            {
                string str = "<root>" + strXML + "</root>";
                XElement elmRoot = XElement.Parse(str);
                foreach (XElement elm in elmRoot.Elements("Item"))
                {
                    if (elm.Attribute("name") !=null)
                    {
                        string value = "";
                        if (elm.Attribute("value") != null)
                            value = elm.Attribute("value").Value;

                        retVal.Add(elm.Attribute("name").Value + "：" + value);
                    }                    
                }
            }
            return string.Join("；", retVal.ToArray());
        }

        /// <summary>
        /// 取得異動相關資料，因為DAL有bug
        /// </summary>
        /// <param name="dataList"></param>
        /// <returns></returns>
        public static Dictionary<string, XElement> GetUpdateRecordInfo(List<string> StudentIDList)
        {
            Dictionary<string, XElement> retVal = new Dictionary<string, XElement>();
            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string strSQL = "select id,context_info from update_record where ref_student_id in(" + string.Join(",",StudentIDList.ToArray()) + ");";
                DataTable dt = qh.Select(strSQL);
                foreach(DataRow dr in dt.Rows)
                {
                    string cont=dr[1].ToString();
                    if(!string.IsNullOrWhiteSpace(cont))
                    {
                        XElement elm =XElement.Parse(cont);
                        retVal.Add(dr[0].ToString(),elm);
                    }
                }
            }
            return retVal;
        }

        /// <summary>
        /// 取得國中領域成績
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string, XElement> GetSemeterDomainScoreByStudentIDList(List<string> StudentIDList)
        {
            Dictionary<string, XElement> retVal = new Dictionary<string, XElement>();
            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string strSQL = "select id,score_info from sems_subj_score where ref_student_id in(" + string.Join(",", StudentIDList.ToArray()) + ");";
                DataTable dt = qh.Select(strSQL);
                foreach (DataRow dr in dt.Rows)
                {
                    string cont = dr[1].ToString();
                    if (!string.IsNullOrWhiteSpace(cont))
                    {
                        StringBuilder sb = new StringBuilder();
                        sb.Append("<root>");
                        sb.Append(cont);
                        sb.Append("</root>");
                        XElement elm = XElement.Parse(sb.ToString());
                        retVal.Add(dr[0].ToString(), elm);
                    }
                }
            }
            return retVal;
        }

        /// <summary>
        ///  取得學期分項學年度學期編號
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string, List<AB_SchoolYearSemesterIdx>> GetSemeterSchoolYearScoreByStudentIDList(List<string> StudentIDList)
        {
            Dictionary<string, List<AB_SchoolYearSemesterIdx>> retVal = new Dictionary<string, List<AB_SchoolYearSemesterIdx>>();
            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string strSQL = "select ref_student_id,school_year,semester,id from sems_entry_score where entry_group=1 and ref_student_id in(" + string.Join(",", StudentIDList.ToArray()) + ");";
                DataTable dt = qh.Select(strSQL);
                foreach (DataRow dr in dt.Rows)
                {
                    string sid = dr[0].ToString();
                    if (!retVal.ContainsKey(sid))
                        retVal.Add(sid, new List<AB_SchoolYearSemesterIdx>());
                    AB_SchoolYearSemesterIdx data = new AB_SchoolYearSemesterIdx ();
                    data.SchoolYear=int.Parse(dr[1].ToString());
                    data.Semester=int.Parse(dr[2].ToString());
                    data.id=dr[3].ToString();
                    retVal[sid].Add(data);
                }
            }
            return retVal;            
        }

        /// <summary>
        /// 取得學期分項成績(不含德行)
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string, XElement> GetSemeterEntryScoreByStudentIDList(List<string> StudentIDList)
        {
            Dictionary<string, XElement> retVal = new Dictionary<string, XElement>();
            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string strSQL = "select id,score_info from sems_entry_score where entry_group=1 and ref_student_id in(" + string.Join(",", StudentIDList.ToArray()) + ");";
                DataTable dt = qh.Select(strSQL);
                foreach (DataRow dr in dt.Rows)
                {
                    string cont = dr[1].ToString();
                    if (!string.IsNullOrWhiteSpace(cont))
                    {
                        XElement elm = XElement.Parse(cont);
                        retVal.Add(dr[0].ToString(), elm);
                    }
                }
            }
            return retVal;
        }

        /// <summary>
        /// 取得高中畢業成績不含德行
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string, XElement> GetGradeScoreSHByStudentIDList(List<string> StudentIDList)
        {
            Dictionary<string, XElement> retVal = new Dictionary<string, XElement>();
            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string strSQL = "select id,grad_score from student where id in(" + string.Join(",", StudentIDList.ToArray()) + ");";
                DataTable dt = qh.Select(strSQL);
                foreach (DataRow dr in dt.Rows)
                {
                    string cont = dr[1].ToString();
                    if (!string.IsNullOrWhiteSpace(cont))
                    {
                        XElement elm = XElement.Parse(cont);
                        retVal.Add(dr[0].ToString(), elm);
                    }
                }
            }
            return retVal;
        }

        /// <summary>
        /// 更新綜合表現紀錄表1(暫時使用，日後需要刪除)
        /// </summary>
        public static void UpdateABQuestions()
        {
            List<UDTQuestionsDataDef> dataList = UDTTransfer.ABUDTQuestionsDataSelectAll();
            
            List<UDTQuestionsDataDef> newDataList = NewAddUDTQuestionsDataDef();
            
            List<string> AddstrList9 = new string[] { "甚佳", "良好", "普通", "欠佳", "甚劣" }.ToList();
            XElement elmRoot = new XElement("Items");
            foreach (string str in AddstrList9)
            {
                XElement elm = new XElement("item");
                elm.SetAttributeValue("key", str);
                elm.SetAttributeValue("has_remark", "False");
                elmRoot.Add(elm);
            }


            foreach (UDTQuestionsDataDef data in dataList)
            {
                if (data.Group == "家庭狀況" && data.Name == "每星期零用錢(元)") data.Name = "每星期零用錢";

                if (data.Group=="學習狀況" && data.Name == "社團幹部") data.QuestionType = "SEMESTER";

                if (data.Group=="學習狀況" && data.Name == "班級幹部") data.QuestionType = "SEMESTER";

                if (data.Group == "自傳" && data.Name == "我國中時的學校生活是") data.Name = "國中時的學校生活";
                if (data.Group == "自傳" && data.Name == "我喜歡的人是") data.Name = "喜歡的人";
                if (data.Group == "自傳" && data.Name == "我最不喜歡做的事是") data.Name = "最不喜歡做的事";
                if (data.Group == "自傳" && data.Name == "我最快樂的回憶") data.Name = "最快樂的回憶";
                if (data.Group == "自傳" && data.Name == "我最要好的朋友是") data.Name = "最要好的朋友";
                if (data.Group == "自傳" && data.Name == "我最喜歡做的事是") data.Name = "最喜歡做的事";
                if (data.Group == "自傳" && data.Name == "我最痛苦的回憶") data.Name = "最痛苦的回憶";
                if (data.Group == "自傳" && data.Name == "家中最了解我的人是") data.Name = "家中最了解我的人";
                if (data.Group == "自傳" && data.Name == "常指導我做功課的人是") data.Name = "常指導我做功課的人";
                if (data.Group == "自傳" && data.Name == "最足以描述自己的幾句話是") data.Name = "最足以描述自己的幾句話";
                if (data.Group == "生活感想" && data.Name == "◎一年來的感想(二年級)") data.Name = "內容1_2";
                if (data.Group == "生活感想" && data.Name == "◎今後努力的目標(二年級)") data.Name = "內容2_2";
                if (data.Group == "生活感想" && data.Name == "◎我對自己的期望(一年級)") data.Name = "內容1_1";
                if (data.Group == "生活感想" && data.Name == "◎為達到理想，我所需要的努力(一年級)") data.Name = "內容2_1";
                if (data.Group == "生活感想" && data.Name == "◎期望師長給予我的幫助(一年級)") data.Name = "內容3_1";
                if (data.Group == "生活感想" && data.Name == "◎期望師長給予我的幫助(二年級)") data.Name = "內容3_3";
                if (data.Group == "生活感想" && data.Name == "內容3_3") data.Name = "內容3_2";
                if (data.Group == "生活感想" && data.Name == "填寫日期(一年級)") data.Name = "填寫日期_1";
                if (data.Group == "生活感想" && data.Name == "填寫日期(二年級)") data.Name = "填寫日期_2";
                if (data.Group == "生活感想" && data.DisplayOrder.Value==4 && data.Name == "填寫日期_1") data.Name = "填寫日期_2";
                if (data.Group == "自我認識" && data.Name == "我的個性(如溫和、急躁、、)") data.Name = "個性_1";
                if (data.Group == "自我認識" && data.Name == "我的優點") data.Name = "優點_1";
                if (data.Group == "自我認識" && data.Name == "我需要改進的地方") data.Name = "需要改進的地方_1";
                if (data.Group == "自我認識" && data.Name == "填寫日期") data.Name = "填寫日期_1";
                if (data.Group == "畢業後計畫" && data.Name == "升學意願(不升學者免填)") data.Name = "升學意願";
                if (data.Group == "畢業後計畫" && data.Name == "希望參加職業訓練種類") data.Name = "參加職業訓練";
                if (data.Group == "畢業後計畫" && data.Name == "受訓地區") data.Name = "受訓地區";
                if (data.Group == "畢業後計畫" && data.Name == "將來職業意願") data.Name = "將來職業";
                if (data.Group == "畢業後計畫" && data.Name == "就業地區") data.Name = "就業地區";
                if (data.Group == "畢業後計畫" && data.Name == "就業意願(升學者免填)") data.Name = "就業意願";
                if (data.Group == "家庭狀況" && data.Name == "直系血親_電話")
                {
                    data.ControlType = "TEXTBOX";
                    data.Items = "";
                }
                if (data.Group == "家庭狀況" && data.Name == "直系血親_存、歿")
                {
                    XElement elmmm = new XElement("Items");
                    XElement elm1 = new XElement("item");
                    elm1.SetAttributeValue("key", "存");
                    elm1.SetAttributeValue("has_remark", "False");
                    elmmm.Add(elm1);
                    XElement elm2 = new XElement("item");
                    elm2.SetAttributeValue("key", "歿");
                    elm2.SetAttributeValue("has_remark", "False");
                    elmmm.Add(elm2);
                    data.Items = elmmm.ToString();
                }

                if (data.Group == "自我認識" || data.Group == "生活感想")
                {
                    data.QuestionType = "SINGLE_ANSWER";
                    data.ControlType = "TEXTBOX";
                }

                // 加入適應情形預設值
                if (data.Group == "適應情形")
                { 
                    // 檢查是否有預設值，沒有加入
                    if (string.IsNullOrWhiteSpace(data.Items))
                        data.Items = elmRoot.ToString();                
                }

                //原住民血統更新
                if (data.Group == "本人概況" && data.Name == "原住民血統")
                {
                    bool change = false;

                    if (data.ControlType != "RADIO_BUTTON")
                    {
                        change = true;
                    }

                    if (string.IsNullOrWhiteSpace(data.Items))
                    {
                        change = true;
                    }

                    if (change)
                    {
                        data.QuestionType = "SINGLE_ANSWER";
                        data.ControlType = "RADIO_BUTTON";
                        data.CanStudentEdit = true;
                        data.CanTeacherEdit = true;
                        data.CanPrint = true;
                        data.Items = "<Items><item key='有' has_remark='True' /><item key='無' has_remark='False' /></Items>";
                    }
                }
            }
         
            // 更新
            UDTTransfer.ABUDTQuestionsDataUpdate(dataList);
            // 新增
            //if (dataList.Count <1000)
            //    UDTTransfer.ABUDTQuestionsDataInsert(newDataList);

            List<string> checkList = new List<string>();
            foreach (UDTQuestionsDataDef data in dataList)
            {
                string key = data.Group + "_" + data.Name;
                if (!checkList.Contains(key))
                    checkList.Add(key);
            }

            List<UDTQuestionsDataDef> addList = new List<UDTQuestionsDataDef>();
            foreach (UDTQuestionsDataDef data in newDataList)
            {
                string key = data.Group + "_" + data.Name;
                if (!checkList.Contains(key))
                    addList.Add(data);
            }

            if(addList.Count > 0)
                UDTTransfer.ABUDTQuestionsDataInsert(addList);

            //List<UDTQuestionsDataDef> sssDataList = new List<UDTQuestionsDataDef>();

            // 單一缺少加入
            //UDTQuestionsDataDef qdadd1 = new UDTQuestionsDataDef();
            //qdadd1.Group = "自傳";
            //qdadd1.Name = "最不喜歡做的事_因為";
            //qdadd1.QuestionType = "SINGLE_ANSWER";
            //qdadd1.ControlType = "TEXTBOX";
            //qdadd1.CanPrint = true;
            //qdadd1.CanStudentEdit = true;
            //qdadd1.CanTeacherEdit = true;

            // 檢查已有不加入
            //bool qdadd1Add = true;
            //foreach (UDTQuestionsDataDef data in dataList)
            //{
            //    if (data.Group == qdadd1.Group && data.Name == qdadd1.Name)
            //    {
            //        qdadd1Add = false;
            //        break;
            //    }
            //}

            //if (qdadd1Add)
            //{
            //    List<UDTQuestionsDataDef> addList = new List<UDTQuestionsDataDef>();
            //    addList.Add(qdadd1);
            //    UDTTransfer.ABUDTQuestionsDataInsert(addList);
            //}

        }

        /// <summary>
        /// 新加入
        /// </summary>
        /// <returns></returns>
        public static List<UDTQuestionsDataDef> NewAddUDTQuestionsDataDef()
        {
            List<UDTQuestionsDataDef> retVal = new List<UDTQuestionsDataDef>();
            List<string> s1 = new List<string>();
            List<string> s2 = new List<string>();
            List<string> s3 = new List<string>();
            //Cloud add
            s1.Add("家中最了解我的人_因為");
            s1.Add("我在家中最怕的人是");
            s1.Add("我在家中最怕的人是_因為");
            s1.Add("我覺得我的優點是");
            s1.Add("我覺得我的缺點是");
            s1.Add("最喜歡的國小（國中）老師");
            s1.Add("最喜歡的國小（國中）老師__因為");
            s1.Add("小學（國中）老師或同學常說我是");
            s1.Add("小學（國中）時我曾在班上登任過的職務有");
            s1.Add("我在小學（國中）得過的獎有");
            s1.Add("我覺得我自己的過去最滿意的是");
            s1.Add("我排遣休閒時間的方法是");
            s1.Add("我最難忘的一件事是");
            s1.Add("自我的心聲_一年級_我目前遇到最大的困難是");
            s1.Add("自我的心聲_一年級_我目前最需要的協助是");
            s1.Add("自我的心聲_二年級_我目前遇到最大的困難是");
            s1.Add("自我的心聲_二年級_我目前最需要的協助是");
            s1.Add("自我的心聲_三年級_我目前遇到最大的困難是");
            s1.Add("自我的心聲_三年級_我目前最需要的協助是");
            //Old
            s1.Add("最喜歡做的事_因為");
            s1.Add("最不喜歡做的事_因為");
            s1.Add("他是怎樣的人");
            s1.Add("喜歡的人_因為");
            s1.Add("讀過且印象最深刻的課外書");
            s1.Add("填寫日期");
            s2.Add("個性_2");
            s2.Add("個性_3");
            s2.Add("填寫日期_2");
            s2.Add("填寫日期_3");
            s2.Add("需要改進的地方_2");
            s2.Add("需要改進的地方_3");
            s2.Add("優點_2");
            s2.Add("優點_3");

            s3.Add("監護人_姓名");
            s3.Add("監護人_性別");
            s3.Add("監護人_關係");
            s3.Add("監護人_電話");
            s3.Add("監護人_通訊地址");

            foreach (string s in s1)
            {
                UDTQuestionsDataDef data = new UDTQuestionsDataDef();
                data.Group = "自傳";
                data.QuestionType = "SINGLE_ANSWER";
                data.ControlType = "TEXTBOX";
                data.Name = s;
                data.CanStudentEdit = true;
                data.CanTeacherEdit = true;
                data.CanPrint = true;
                retVal.Add(data);
            }
            foreach (string s in s2)
            {
                UDTQuestionsDataDef data = new UDTQuestionsDataDef();
                data.Group = "自我認識";
                data.QuestionType = "SINGLE_ANSWER";
                data.ControlType = "TEXTBOX";
                data.Name = s;
                data.CanStudentEdit = true;
                data.CanTeacherEdit = true;
                data.CanPrint = true;
                retVal.Add(data);
            }

            foreach (string s in s3)
            {
                UDTQuestionsDataDef data = new UDTQuestionsDataDef();
                data.Group = "家庭狀況";
                data.QuestionType = "SINGLE_ANSWER";
                data.ControlType = "TEXTBOX";
                data.Name = s;
                data.CanStudentEdit = true;
                data.CanTeacherEdit = true;
                data.CanPrint = true;
                retVal.Add(data);
            }

            UDTQuestionsDataDef d1 = new UDTQuestionsDataDef();
            d1.Group = "家庭狀況";
            d1.QuestionType = "Relative";
            d1.ControlType = "GRID_COMBOBOX";
            d1.Name = "直系血親_電話";
            d1.CanStudentEdit = true;
            d1.CanTeacherEdit = true;
            d1.CanPrint = true;
            retVal.Add(d1);

            UDTQuestionsDataDef d2 = new UDTQuestionsDataDef();
            d2.Group = "家庭狀況";
            d2.QuestionType = "SINGLE_ANSWER";
            d2.ControlType = "TEXTBOX";
            d2.Name = "兄弟姊妹_排行";
            d2.CanStudentEdit = true;
            d2.CanTeacherEdit = true;
            d2.CanPrint = true;
            retVal.Add(d2);

            UDTQuestionsDataDef d3 = new UDTQuestionsDataDef();
            d3.Group = "本人概況";
            d3.QuestionType = "SINGLE_ANSWER";
            d3.ControlType = "RADIO_BUTTON";
            d3.Name = "原住民血統";
            d3.CanStudentEdit = true;
            d3.CanTeacherEdit = true;
            d3.CanPrint = true;
            d3.Items = "<Items><item key='有' has_remark='True' /><item key='無' has_remark='False' /></Items>";
            retVal.Add(d3);

            UDTQuestionsDataDef d4 = new UDTQuestionsDataDef();
            d4.Group = "家庭狀況";
            d4.QuestionType = "Relative";
            d4.ControlType = "GRID_TEXTBOX";
            d4.Name = "直系血親_原國籍";
            d4.CanStudentEdit = true;
            d4.CanTeacherEdit = true;
            d4.CanPrint = true;
            retVal.Add(d4);

            return retVal;
        }

        /// <summary>
        /// 透過教師ID取得晤談紀錄簽認表用資料
        /// </summary>
        /// <param name="TeacherIDList"></param>
        /// <returns></returns>
        public static Dictionary<string,List<Rpt_InterviewRecord>> GetInterviewRecordDictByTeacherID1(List<string> TeacherIDList)
        {
            Dictionary<string,List<Rpt_InterviewRecord>> retVal = new Dictionary<string,List<Rpt_InterviewRecord>> ();
            if (TeacherIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string qry = "select teacher.id as tid,student.id as sid,class.class_name,student.seat_no,student.name,teacher.teacher_name,teacher.nickname,interview_type,interviewee_type,interview_date,content_digest from $counsel.interview_record inner join student on  $counsel.interview_record.ref_student_id=student.id left join class on student.ref_class_id=class.id inner join teacher on teacher.id= $counsel.interview_record.teacher_id where teacher.id in("+string.Join(",",TeacherIDList.ToArray())+") order by class.class_name,seat_no,interview_date;";
                DataTable dt = qh.Select(qry);
                foreach (DataRow dr in dt.Rows)
                {
                    string tid=dr[0].ToString();
                    if(!retVal.ContainsKey(tid))
                        retVal.Add(tid,new List<Rpt_InterviewRecord>());

                    Rpt_InterviewRecord data = new Rpt_InterviewRecord ();
                    data.TeacherID=tid;
                    data.StudentID=dr[1].ToString();
                    data.ClassName=dr["class_name"].ToString();
                    string seatNo=dr["seat_no"].ToString();
                    if (!string.IsNullOrWhiteSpace(seatNo))
                        data.SeatNo = int.Parse(seatNo);
                    data.TeacherName=dr["teacher_name"].ToString();
                    string tniname = dr["nickname"].ToString();
                    if (!string.IsNullOrWhiteSpace(tniname))
                        data.TeacherName += "(" + tniname + ")";

                    data.Name = dr["name"].ToString();
                    data.interview_type = dr["interview_type"].ToString();
                    data.interviewee_type = dr["interviewee_type"].ToString();
                    data.interview_date = DateTime.Parse(dr["interview_date"].ToString());
                    data.content_digest = dr["content_digest"].ToString();
                    retVal[tid].Add(data);
                }
            }
            return retVal;
        }
    }
}
