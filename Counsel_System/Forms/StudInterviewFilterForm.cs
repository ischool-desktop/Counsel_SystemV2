﻿using System;
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
using System.IO;
using System.Xml;
using System.Xml.Linq;

namespace Counsel_System2.Forms
{
    public partial class StudInterviewFilterForm : FISCA.Presentation.Controls.BaseForm
    {
        private List<string> _StudentIDList = new List<string>();

        public StudInterviewFilterForm(List<string> studentIDList)
        {
            _StudentIDList = studentIDList;

            InitializeComponent();

            List<string> counselTypeKindList = new List<string>();
            counselTypeKindList.Add("家人議題");
            counselTypeKindList.Add("違規行為");
            counselTypeKindList.Add("心理困擾");
            counselTypeKindList.Add("學習問題");
            counselTypeKindList.Add("性別議題");
            counselTypeKindList.Add("人際關係");
            counselTypeKindList.Add("生涯規劃");
            counselTypeKindList.Add("自傷/自殺");
            counselTypeKindList.Add("生活適應");
            counselTypeKindList.Add("生活作息/常規");
            counselTypeKindList.Add("家長期望");
            counselTypeKindList.Add("健康問題");
            counselTypeKindList.Add("情緒不穩");
            counselTypeKindList.Add("法定通報-兒少保護");
            counselTypeKindList.Add("法定通報-高風險家庭");
            counselTypeKindList.Add("法定通報-家暴(18 歲以上)");
            counselTypeKindList.Add("其他(含生活關懷)");

            foreach (var item in counselTypeKindList)
            {
                listViewEx1.Items.Add(item);
            }


            List<string> authorRoleList = new List<string>();
            authorRoleList.Add("輔導老師");
            authorRoleList.Add("認輔老師");
            authorRoleList.Add("班導師");

            foreach (var authorRole in authorRoleList)
            {
                comboBoxEx1.Items.Add(authorRole);
            }
        }



        private void buttonX1_Click(object sender, EventArgs e)
        {
            var isFilterSchoolyearSemester = checkBox1.Checked;
            var isFilterDayBetween = checkBox2.Checked;
            var isFilterAuthorRole = checkBox3.Checked;
            var isFilterCounselTypeKind = checkBox4.Checked;

            var filterSchoolyear = textBoxX1.Text;
            var filterSemester = textBoxX2.Text;
            var filterDateBegin = dateTimeInput1.Value;
            var filterDateEnd = dateTimeInput2.Value.AddDays(1);
            var filterAuthorRole = comboBoxEx1.Text;

            var filterCounselTypeKindFilterList = new List<string>();
            foreach (ListViewItem item in listViewEx1.Items)
            {
                if (item.Checked)
                {
                    filterCounselTypeKindFilterList.Add(item.Text);
                }
            }

            var showAttendeesDetial = checkBox5.Checked;
            var showCounselTypeDetial = checkBox6.Checked;
            var showCounselTypeKindDetial = checkBox7.Checked;

            Workbook wb = new Workbook();

            BackgroundWorker printWorker = new BackgroundWorker() { WorkerReportsProgress = true };
            printWorker.DoWork += delegate
            {
                //整理學生StudentRecord 資料
                Dictionary<string, StudentRecord> dicStudentRecord = new Dictionary<string, StudentRecord>();

                foreach (var studentRecord in K12.Data.Student.SelectByIDs(_StudentIDList))
                {
                    dicStudentRecord.Add(studentRecord.ID, studentRecord);
                }
                printWorker.ReportProgress(5);
                #region 篩選晤談紀錄
                FISCA.UDT.AccessHelper _AccessHelper = new FISCA.UDT.AccessHelper();
                var interviewRecordList = _AccessHelper.Select<DAO.UDT_CounselStudentInterviewRecordDef>("ref_student_id in (" + string.Join(",", _StudentIDList) + ")");
                Dictionary<string, List<DAO.UDT_CounselStudentInterviewRecordDef>> dicStudentInterviewRecord = new Dictionary<string, List<DAO.UDT_CounselStudentInterviewRecordDef>>();
                printWorker.ReportProgress(20);

                int progressCount = 0;
                //整理晤談資料，符合條件才加入
                foreach (var interviewRecord in interviewRecordList)
                {
                    //學年度 學期篩選
                    if (isFilterSchoolyearSemester)
                    {
                        if (interviewRecord.schoolyear != filterSchoolyear || interviewRecord.semester != filterSemester)
                        {
                            continue;
                        }
                    }

                    //日期區間篩選
                    if (isFilterDayBetween)
                    {
                        //2017/1/20 穎驊更新，發現使用者有些晤談紀錄會沒有晤談日期，會造成interviewRecord.InterviewDate.Value 其中的 InterviewDate 變成 nullReference 的狀況爆掉
                        //因此在這邊處理，若沒有晤談資料沒有晤談日期又有勾選日期區間篩選，因為無法比較日期，直接排除掉此晤談資料

                        if (interviewRecord.InterviewDate == null)
                        {
                            continue;
                        }
                        else 
                        {
                            if (interviewRecord.InterviewDate.Value < filterDateBegin ||
                                interviewRecord.InterviewDate >= filterDateEnd)
                            {
                                continue;
                            }
                        
                        }
                        
                    }

                    //記錄人篩選
                    if (isFilterAuthorRole)
                    {
                        if (interviewRecord.authorRole != filterAuthorRole)
                        {
                            continue;
                        }
                    }

                    //案件類別篩選
                    if (isFilterCounselTypeKind)
                    {
                        XmlDocument doc3 = new XmlDocument();
                        //幫忙加根目錄
                        string xmlContent3 = "<root>" + interviewRecord.CounselTypeKind + "</root>";
                        doc3.LoadXml(xmlContent3);
                        XmlNode newNode3 = doc3.DocumentElement;
                        doc3.AppendChild(newNode3);
                        XElement xmlabs3 = XElement.Parse(doc3.OuterXml);
                        bool pass = false;

                        foreach (XElement abs in xmlabs3.Elements("Item"))
                        {
                            if (filterCounselTypeKindFilterList.Contains(abs.Attribute("name").Value))
                            {
                                pass = true;
                            }
                        }
                        if (!pass)
                        {
                            continue;
                        }

                    }

                    if (!dicStudentInterviewRecord.ContainsKey("" + interviewRecord.StudentID))
                    {
                        dicStudentInterviewRecord.Add("" + interviewRecord.StudentID, new List<DAO.UDT_CounselStudentInterviewRecordDef>());

                        dicStudentInterviewRecord["" + interviewRecord.StudentID].Add(interviewRecord);

                    }
                    else
                    {
                        dicStudentInterviewRecord["" + interviewRecord.StudentID].Add(interviewRecord);
                    }

                    progressCount++;

                    printWorker.ReportProgress(20 + 50 * progressCount / interviewRecordList.Count);
                }
                #endregion

                int row_index = 1;

                DataTable dt = new DataTable();
                var basicDataTitleList = new List<string>();
                var attendeesDataTitleList = new List<string>();
                var counselTypeDataTitleList = new List<string>();
                var counselTypeKindDataTitleList = new List<string>();
                #region 基本欄位
                //所有Datatable 欄位--單項目
                dt.Columns.Add("學號");
                dt.Columns.Add("學生姓名");
                dt.Columns.Add("性別");
                dt.Columns.Add("班級");
                dt.Columns.Add("座號");
                dt.Columns.Add("晤談編號");
                dt.Columns.Add("學年度");
                dt.Columns.Add("學期");
                dt.Columns.Add("晤談日期");
                dt.Columns.Add("晤談時間");
                dt.Columns.Add("晤談動機");
                dt.Columns.Add("輔導對象");
                dt.Columns.Add("晤談方式");
                dt.Columns.Add("地點");
                dt.Columns.Add("內容要點");
                dt.Columns.Add("記錄人的登入帳號");
                dt.Columns.Add("記錄人的姓名");
                dt.Columns.Add("是否公開");
                dt.Columns.Add("記錄人");

                //所有Datatable 欄位--多項目
                dt.Columns.Add("參與人員");
                dt.Columns.Add("處理方式");
                dt.Columns.Add("案件類別");

                //基本資料Title--單項目
                basicDataTitleList.Add("學號");
                basicDataTitleList.Add("學生姓名");
                basicDataTitleList.Add("性別");
                basicDataTitleList.Add("班級");
                basicDataTitleList.Add("座號");
                basicDataTitleList.Add("晤談編號");
                basicDataTitleList.Add("學年度");
                basicDataTitleList.Add("學期");
                basicDataTitleList.Add("晤談日期");
                basicDataTitleList.Add("晤談時間");
                basicDataTitleList.Add("晤談動機");
                basicDataTitleList.Add("輔導對象");
                basicDataTitleList.Add("晤談方式");
                basicDataTitleList.Add("地點");

                //基本資料Title--多項目
                basicDataTitleList.Add("參與人員");
                basicDataTitleList.Add("處理方式");
                basicDataTitleList.Add("案件類別");

                basicDataTitleList.Add("內容要點");
                //basicDataTitleList.Add("記錄人的登入帳號");
                basicDataTitleList.Add("記錄人的姓名");
                basicDataTitleList.Add("是否公開");
                basicDataTitleList.Add("記錄人");
                #endregion

                progressCount = 0;
                //開始整理資料
                foreach (var student_id in _StudentIDList)
                {
                    #region 整理學生的紀錄
                    StudentRecord stuRec = dicStudentRecord[student_id];
                    if (dicStudentInterviewRecord.ContainsKey(student_id))
                    {
                        foreach (var InterviewRecord in dicStudentInterviewRecord[student_id])
                        {
                            //每有一筆晤談紀錄 就有一個新Row
                            DataRow row = dt.NewRow();
                            //學號
                            row["學號"] = stuRec.StudentNumber;
                            //學生姓名
                            row["學生姓名"] = stuRec.Name;
                            //性別
                            row["性別"] = stuRec.Gender;
                            //班級
                            row["班級"] = (stuRec.Class != null ? stuRec.Class.Name : "");
                            //座號
                            row["座號"] = stuRec.SeatNo;
                            //晤談編號
                            row["晤談編號"] = InterviewRecord.InterviewNo;
                            //學年度
                            row["學年度"] = InterviewRecord.schoolyear;
                            //學期
                            row["學期"] = InterviewRecord.semester;


                            //2017/1/20 穎驊更新，發現使用者有些晤談紀錄會沒有晤談日期，會造成interviewRecord.InterviewDate.Value 其中的 InterviewDate 變成 nullReference 的狀況爆掉
                            // 因此做了例外處理，有日期則顯示日期，沒有則顯示空字串
                            //晤談日期
                            row["晤談日期"] = InterviewRecord.InterviewDate.HasValue? InterviewRecord.InterviewDate.Value.ToShortDateString():"";


                            //晤談時間
                            row["晤談時間"] = InterviewRecord.InterviewTime;
                            //晤談動機
                            row["晤談動機"] = InterviewRecord.Cause;
                            //輔導對象
                            row["輔導對象"] = InterviewRecord.IntervieweeType;
                            //晤談方式
                            row["晤談方式"] = InterviewRecord.InterviewType;
                            //地點
                            row["地點"] = InterviewRecord.Place;
                            //內容要點
                            row["內容要點"] = InterviewRecord.ContentDigest;
                            //記錄人的登入帳號
                            row["記錄人的登入帳號"] = InterviewRecord.AuthorID;
                            //記錄人的姓名
                            row["記錄人的姓名"] = InterviewRecord.AuthorName;
                            //是否公開
                            row["是否公開"] = (InterviewRecord.isPublic ? "是" : "否");
                            //記錄人
                            row["記錄人"] = InterviewRecord.authorRole;

                            #region row["參與人員"] 整理
                            XmlDocument doc1 = new XmlDocument();
                            //幫忙加根目錄
                            string xmlContent1 = "<root>" + InterviewRecord.Attendees + "</root>";
                            doc1.LoadXml(xmlContent1);
                            XmlNode newNode1 = doc1.DocumentElement;
                            doc1.AppendChild(newNode1);
                            XElement xmlabs1 = XElement.Parse(doc1.OuterXml);
                            string attendees = "";
                            string attendees_for_basic = "";

                            foreach (XElement abs in xmlabs1.Elements("Item"))
                            {
                                attendees = abs.Attribute("name").Value;

                                attendees_for_basic += abs.Attribute("name").Value;
                                if (abs != xmlabs1.LastNode)
                                {
                                    attendees_for_basic += "、";
                                }

                                //假如使用者有需要參與人員分析資料，則新增欄位，並在其值填 "是"
                                if (!dt.Columns.Contains("參與人員:" + attendees))
                                {
                                    dt.Columns.Add("參與人員:" + attendees);
                                    attendeesDataTitleList.Add("參與人員:" + attendees);
                                }
                                //參與人員
                                row["參與人員"] = attendees_for_basic;
                                row["參與人員:" + attendees] = "是";


                            }
                            #endregion

                            #region row["處理方式"]整理
                            XmlDocument doc2 = new XmlDocument();
                            //幫忙加根目錄
                            string xmlContent2 = "<root>" + InterviewRecord.CounselType + "</root>";
                            doc2.LoadXml(xmlContent2);
                            XmlNode newNode2 = doc2.DocumentElement;
                            doc2.AppendChild(newNode2);
                            XElement xmlabs2 = XElement.Parse(doc2.OuterXml);
                            string CounselType = "";
                            string CounselType_for_basic = "";

                            foreach (XElement abs in xmlabs2.Elements("Item"))
                            {

                                CounselType_for_basic += abs.Attribute("name").Value;
                                if (abs.Attribute("name").Value == "其他")
                                {
                                    if (abs.Attribute("remark") != null)
                                        CounselType_for_basic += ":" + abs.Attribute("remark").Value;
                                }
                                if (abs != xmlabs2.LastNode)
                                {
                                    CounselType_for_basic += "、";
                                }


                                CounselType = abs.Attribute("name").Value;
                                if (abs.Attribute("name").Value == "其他")
                                {
                                    if (abs.Attribute("remark") != null)
                                        CounselType += ":" + abs.Attribute("remark").Value;
                                }
                                //假如使用者有需要處理分析資料，則新增欄位，並在其值填 "是"
                                if (!dt.Columns.Contains("處理方式:" + CounselType))
                                {
                                    counselTypeDataTitleList.Add("處理方式:" + CounselType);
                                    dt.Columns.Add("處理方式:" + CounselType);
                                }



                                //處理方式
                                row["處理方式"] = CounselType_for_basic;
                                row["處理方式:" + CounselType] = "是";
                            }
                            #endregion

                            #region row["案件類別"]整理
                            XmlDocument doc3 = new XmlDocument();
                            //幫忙加根目錄
                            string xmlContent3 = "<root>" + InterviewRecord.CounselTypeKind + "</root>";
                            doc3.LoadXml(xmlContent3);
                            XmlNode newNode3 = doc3.DocumentElement;
                            doc3.AppendChild(newNode3);
                            XElement xmlabs3 = XElement.Parse(doc3.OuterXml);
                            string CounselTypeKind = "";
                            string CounselTypeKind_for_basic = "";

                            foreach (XElement abs in xmlabs3.Elements("Item"))
                            {
                                CounselTypeKind_for_basic += abs.Attribute("name").Value;
                                if (abs.Attribute("name").Value == "其他")
                                {
                                    if (abs.Attribute("remark") != null)
                                        CounselTypeKind_for_basic += ":" + abs.Attribute("remark").Value;
                                }
                                if (abs != xmlabs3.LastNode)
                                {
                                    CounselTypeKind_for_basic += "、";
                                }


                                CounselTypeKind = abs.Attribute("name").Value;
                                if (abs.Attribute("name").Value == "其他")
                                {
                                    if (abs.Attribute("remark") != null)
                                        CounselTypeKind += ":" + abs.Attribute("remark").Value;
                                }
                                //假如使用者有需要案件類別分析資料，則新增欄位，並在其值填 "是"
                                if (!dt.Columns.Contains("案件類別:" + CounselTypeKind))
                                {
                                    dt.Columns.Add("案件類別:" + CounselTypeKind);
                                    counselTypeKindDataTitleList.Add("案件類別:" + CounselTypeKind);
                                }
                                //案件類別
                                row["案件類別"] = CounselTypeKind_for_basic;
                                row["案件類別:" + CounselTypeKind] = "是";

                            }
                            #endregion

                            dt.Rows.Add(row);
                        }
                    }
                    #endregion
                    progressCount++;
                    printWorker.ReportProgress(70 + 25 * progressCount / _StudentIDList.Count);
                }

                int col_index = 0;

                #region 加入Excel 表單 每欄Title

                wb.Open(new MemoryStream(Properties.Resources.學生晤談紀錄篩選), FileFormatType.Excel2003);
                Cells cs = wb.Worksheets[0].Cells;

                //基本資料
                foreach (string title in basicDataTitleList)
                {
                    cs[0, col_index].PutValue(title);

                    col_index++;
                }

                //參與人員分析資料
                if (showAttendeesDetial)
                {
                    foreach (string title in attendeesDataTitleList)
                    {
                        cs[0, col_index].PutValue(title);

                        col_index++;
                    }
                }
                //處理方式分析資料
                if (showCounselTypeDetial)
                {
                    foreach (string title in counselTypeDataTitleList)
                    {
                        cs[0, col_index].PutValue(title);

                        col_index++;
                    }
                }
                //案件類別分析資料
                if (showCounselTypeKindDetial)
                {
                    foreach (string title in counselTypeKindDataTitleList)
                    {
                        cs[0, col_index].PutValue(title);

                        col_index++;
                    }
                }
                #endregion

                #region 開始Excel填值
                foreach (DataRow row in dt.Rows)
                {
                    col_index = 0;

                    //基本資料
                    foreach (string title in basicDataTitleList)
                    {
                        cs[row_index, col_index].PutValue(row[title]);

                        col_index++;
                    }

                    //參與人員分析資料
                    if (showAttendeesDetial)
                    {
                        foreach (string title in attendeesDataTitleList)
                        {
                            cs[row_index, col_index].PutValue(row[title]);

                            col_index++;
                        }
                    }

                    //處理方式分析資料
                    if (showCounselTypeDetial)
                    {
                        foreach (string title in counselTypeDataTitleList)
                        {
                            cs[row_index, col_index].PutValue(row[title]);

                            col_index++;
                        }
                    }


                    //案件類別分析資料
                    if (showCounselTypeKindDetial)
                    {
                        foreach (string title in counselTypeKindDataTitleList)
                        {
                            cs[row_index, col_index].PutValue(row[title]);

                            col_index++;
                        }
                    }

                    row_index++;


                }
                #endregion

                #region 舊列印方式
                //foreach (var student_id in student_id_List) 
                //{
                //    StudentRecord stuRec = new StudentRecord();

                //    if (InterviewRecord_Dict.ContainsKey(student_id))
                //    {
                //        if (studentRecord_Dict.ContainsKey(student_id)) 
                //        {
                //            stuRec = studentRecord_Dict[student_id];
                //        }

                //        foreach (var InterviewRecord in InterviewRecord_Dict[student_id]) 
                //        {
                //            //學生姓名
                //            cs[row_index, 0].PutValue(stuRec.Name);
                //            //學號
                //            cs[row_index, 1].PutValue(stuRec.StudentNumber);
                //            //性別
                //            cs[row_index, 2].PutValue(stuRec.Gender);
                //            //班級
                //            cs[row_index, 3].PutValue(  stuRec.Class !=null? stuRec.Class.Name:"");
                //            //座號
                //            cs[row_index, 4].PutValue(stuRec.SeatNo);
                //            //晤談編號
                //            cs[row_index, 5].PutValue(InterviewRecord.InterviewNo);
                //            //學年度
                //            cs[row_index, 6].PutValue(InterviewRecord.schoolyear);
                //            //學期
                //            cs[row_index, 7].PutValue(InterviewRecord.semester);
                //            //晤談日期
                //            cs[row_index, 8].PutValue(InterviewRecord.InterviewDate.Value.ToShortDateString());
                //            //晤談時間
                //            cs[row_index, 9].PutValue(InterviewRecord.InterviewTime);
                //            //晤談動機
                //            cs[row_index, 10].PutValue(InterviewRecord.Cause);
                //            //輔導對象
                //            cs[row_index, 11].PutValue(InterviewRecord.IntervieweeType);
                //            //晤談方式
                //            cs[row_index, 12].PutValue(InterviewRecord.InterviewType);
                //            //地點
                //            cs[row_index, 13].PutValue(InterviewRecord.Place);

                //            XmlDocument doc1 = new XmlDocument();
                //            //幫忙加根目錄
                //            string xmlContent1 = "<root>" + InterviewRecord.Attendees + "</root>";
                //            doc1.LoadXml(xmlContent1);
                //            XmlNode newNode1 = doc1.DocumentElement;
                //            doc1.AppendChild(newNode1);
                //            XElement xmlabs1 = XElement.Parse(doc1.OuterXml);
                //            string attendees = "";

                //            foreach (XElement abs in xmlabs1.Elements("Item"))
                //            {
                //              attendees += abs.Attribute("name").Value;
                //              if (abs != xmlabs1.LastNode) 
                //              {
                //                  attendees += "、";                          
                //              }                         
                //            }

                //            //參與人員
                //            cs[row_index, 14].PutValue(attendees);

                //            XmlDocument doc2 = new XmlDocument();
                //            //幫忙加根目錄
                //            string xmlContent2 = "<root>" + InterviewRecord.CounselType + "</root>";
                //            doc2.LoadXml(xmlContent2);
                //            XmlNode newNode2 = doc2.DocumentElement;
                //            doc2.AppendChild(newNode2);
                //            XElement xmlabs2 = XElement.Parse(doc2.OuterXml);
                //            string CounselType = "";

                //            foreach (XElement abs in xmlabs2.Elements("Item"))
                //            {
                //                CounselType += abs.Attribute("name").Value;
                //                if (abs.Attribute("name").Value == "其他") 
                //                {
                //                    if(abs.Attribute("remark")!=null)
                //                    CounselType += ":"+abs.Attribute("remark").Value;                            
                //                }
                //                if (abs != xmlabs2.LastNode)
                //                {
                //                    CounselType += "、";
                //                }
                //            }

                //            //處理方式
                //            cs[row_index, 15].PutValue(CounselType);
                //            XmlDocument doc3 = new XmlDocument();
                //            //幫忙加根目錄
                //            string xmlContent3 = "<root>" + InterviewRecord.CounselTypeKind + "</root>";
                //            doc3.LoadXml(xmlContent3);
                //            XmlNode newNode3 = doc3.DocumentElement;
                //            doc3.AppendChild(newNode3);
                //            XElement xmlabs3 = XElement.Parse(doc3.OuterXml);
                //            string CounselTypeKind = "";
                //            foreach (XElement abs in xmlabs3.Elements("Item"))
                //            {
                //                CounselTypeKind += abs.Attribute("name").Value;
                //                if (abs.Attribute("name").Value == "其他")
                //                {
                //                    if (abs.Attribute("remark") != null)
                //                        CounselTypeKind += ":" + abs.Attribute("remark").Value;
                //                }
                //                if (abs != xmlabs3.LastNode)
                //                {
                //                    CounselTypeKind += "、";
                //                }
                //            }

                //            //案件類別
                //            cs[row_index, 16].PutValue(CounselTypeKind);

                //            //參與人員(XML)
                //            //cs[row_index, 14].PutValue(InterviewRecord.Attendees);
                //            //處理方式(XML)
                //            //cs[row_index, 15].PutValue(InterviewRecord.CounselType);
                //            //案件類別(XML)
                //            //cs[row_index, 16].PutValue(InterviewRecord.CounselTypeKind);

                //            //內容要點
                //            cs[row_index, 17].PutValue(InterviewRecord.ContentDigest);
                //            //記錄人的登入帳號
                //            cs[row_index, 18].PutValue(InterviewRecord.AuthorID);
                //            //記錄人的姓名
                //            cs[row_index, 19].PutValue(InterviewRecord.AuthorName);
                //            //是否公開(認輔老師、班導師之間互相可見)
                //            cs[row_index, 20].PutValue(InterviewRecord.isPublic? "是":"否");
                //            //記錄人
                //            cs[row_index, 21].PutValue(InterviewRecord.authorRole);                       
                //            row_index++;

                //        }


                //    }
                //    else 
                //    {
                //        if (studentRecord_Dict.ContainsKey(student_id))
                //        {
                //            stuRec = studentRecord_Dict[student_id];
                //        }

                //    }          
                //} 
                #endregion

                printWorker.ReportProgress(100);
            };
            printWorker.RunWorkerCompleted += delegate
            {
                // 以後記得存Excel 都用新版的Xlsx，可以避免ㄧ些不必要的問題(EX: sheet 只能到1023張)
                SaveFileDialog save = new SaveFileDialog();
                save.Title = "另存新檔";
                save.FileName = "學生晤談紀錄篩選";
                save.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";

                if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        wb.Save(save.FileName, FileFormatType.Excel2003);
                        System.Diagnostics.Process.Start(save.FileName);


                    }
                    catch
                    {
                        MessageBox.Show("檔案儲存失敗");


                    }
                }
            };
            printWorker.ProgressChanged += delegate (object s, ProgressChangedEventArgs e2)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("匯出學生晤談紀錄", e2.ProgressPercentage);
            };
            printWorker.RunWorkerAsync();

            this.Close();
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        //使用者輸入學年學期後　自動勾選
        private void textBoxX1_TextChanged(object sender, EventArgs e)
        {
            checkBox1.Checked = true;
        }

        //使用者輸入時間區段後　自動勾選
        private void dateTimeInput1_TextChanged(object sender, EventArgs e)
        {
            checkBox2.Checked = true;
        }

        //使用者選擇記錄人　自動勾選
        private void comboBoxEx1_SelectedIndexChanged(object sender, EventArgs e)
        {
            checkBox3.Checked = true;
        }

        //使用者勾選案件類別　自動勾選
        private void listViewEx1_ItemChecked(object sender, ItemCheckedEventArgs e)
        {

            if (e.Item.Checked == true)
            {
                checkBox4.Checked = true;

            }
            else
            {
                int check_counter = 0;

                foreach (ListViewItem item in listViewEx1.Items)
                {
                    if (item.Checked)
                    {
                        check_counter++;
                    }
                }

                if (check_counter == 0)
                {
                    checkBox4.Checked = false;
                }
            }
        }




    }
}
