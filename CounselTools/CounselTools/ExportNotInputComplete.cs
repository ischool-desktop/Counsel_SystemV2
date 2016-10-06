using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Cells;
using System.ComponentModel;
using System.Text;
using System.Data;
using System.IO;
using System.Drawing;


namespace CounselTools
{
    public class ExportNotInputABCardComplete
    {
        List<string> _StudentIDList;
        BackgroundWorker _bgLoadData;

        //2106/9/20 穎驊新增，用來存放所有學生詳細資料之Dict
        Dictionary<string,Dictionary<string,string>> StuData = new Dictionary<string,Dictionary<string,string>>();

        public ExportNotInputABCardComplete(List<string> StudentIDList)
        {
            _bgLoadData = new BackgroundWorker();
            _bgLoadData.DoWork += _bgLoadData_DoWork;
            _bgLoadData.ProgressChanged += _bgLoadData_ProgressChanged;
            _bgLoadData.WorkerReportsProgress = true;
            _bgLoadData.RunWorkerCompleted += _bgLoadData_RunWorkerCompleted;
            // 學生編號
            _StudentIDList = StudentIDList;

            // 載入資料
            _bgLoadData.RunWorkerAsync();

        }

        void _bgLoadData_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("綜合紀錄表輸入進度表產生中", e.ProgressPercentage);
        }

        void _bgLoadData_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                FISCA.Presentation.Controls.MsgBox.Show("產生過程發生錯誤," + e.Error.Message);
            }
            else
            {
                try
                {
                    Workbook wb = e.Result as Workbook;
                    if (wb != null)
                    {
                        Utility.CompletedXlsx("輔導綜合紀錄表輸入進度表", wb);
                    }
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("產生 Excel 失敗," + ex.Message);
                }
            }

        }

        void _bgLoadData_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                _bgLoadData.ReportProgress(5);
                Gobal._multiple_recordDict.Clear();
                Gobal._priority_dataDict.Clear();
                Gobal._relativeDict.Clear();
                Gobal._semester_dataDict.Clear();
                Gobal._siblingDict.Clear();
                Gobal._single_recordDict.Clear();
                Gobal._yearly_dataDict.Clear();

                Gobal._multiple_recordDict = Utility.GetABCardAnswerDataByStudentIDList(_StudentIDList, "multiple_record");
                _bgLoadData.ReportProgress(10);
                Gobal._priority_dataDict = Utility.GetABCardAnswerDataByStudentIDList(_StudentIDList, "priority_data");
                _bgLoadData.ReportProgress(15);
                Gobal._relativeDict = Utility.GetABCardAnswerDataByStudentIDList(_StudentIDList, "relative");
                _bgLoadData.ReportProgress(20);
                Gobal._semester_dataDict = Utility.GetABCardAnswerDataByStudentIDList(_StudentIDList, "semester_data");
                _bgLoadData.ReportProgress(25);
                Gobal._siblingDict = Utility.GetABCardAnswerDataByStudentIDList(_StudentIDList, "sibling");
                _bgLoadData.ReportProgress(30);
                Gobal._single_recordDict = Utility.GetABCardAnswerDataByStudentIDList(_StudentIDList, "single_record");
                _bgLoadData.ReportProgress(35);
                Gobal._yearly_dataDict = Utility.GetABCardAnswerDataByStudentIDList(_StudentIDList, "yearly_data");
                _bgLoadData.ReportProgress(50);
                // 取得學生資料
                List<ClassStudent> ClassStudents = Utility.GetClassStudentByStudentIDList(_StudentIDList);
                _bgLoadData.ReportProgress(65);

                List<string> cpGNameList = new List<string>();

                cpGNameList.Add("個人資料");

                //2016/9/8 穎驊新增欄位
                cpGNameList.Add("監護人資料");
                cpGNameList.Add("尊親屬資料");
                cpGNameList.Add("兄弟姊妹資料");
                cpGNameList.Add("身高及體重");
                cpGNameList.Add("家庭訊息");
                cpGNameList.Add("學習");
                cpGNameList.Add("幹部資訊");
                cpGNameList.Add("自我認識");
                cpGNameList.Add("生活感想");
                cpGNameList.Add("畢業後規劃");
                cpGNameList.Add("自傳");

                foreach (ClassStudent cs in ClassStudents)
                {
                    Dictionary<string, ICheckProcess> CheckProcDict = new Dictionary<string, ICheckProcess>();
                    // 開始檢查
                    foreach (string cpGName in cpGNameList)
                    {
                        switch (cpGName)
                        {
                            case "個人資料":
                                CheckProcDict.Add(cpGName, new CheckProcess1());
                                break;

                            case "監護人資料":
                                CheckProcDict.Add(cpGName, new CheckProcess11());
                                break;

                            case "尊親屬資料":
                                CheckProcDict.Add(cpGName, new CheckProcess13());
                                break;

                            case "兄弟姊妹資料":
                                CheckProcDict.Add(cpGName, new CheckProcess15());
                                break;

                            case "身高及體重":
                                CheckProcDict.Add(cpGName, new CheckProcess10());
                                break;

                            case "家庭訊息":
                                CheckProcDict.Add(cpGName, new CheckProcess14());
                                break;

                            case "學習":
                                CheckProcDict.Add(cpGName, new CheckProcess3());
                                break;

                            case "幹部資訊":
                                CheckProcDict.Add(cpGName, new CheckProcess12());
                                break;

                            case "自我認識":
                                CheckProcDict.Add(cpGName, new CheckProcess5());
                                break;
                            case "生活感想":
                                CheckProcDict.Add(cpGName, new CheckProcess6());
                                break;

                            case "畢業後規劃":
                                CheckProcDict.Add(cpGName, new CheckProcess7());
                                break;

                            case "自傳":
                                CheckProcDict.Add(cpGName, new CheckProcess4());
                                break;

                        }
                    }

                    foreach (string cpGName in CheckProcDict.Keys)
                    {
                        ICheckProcess cp = CheckProcDict[cpGName] as ICheckProcess;
                        if (cp != null)
                        {
                            cp.SetGroupName(cpGName);
                            cp.SetStudent(cs);
                            cp.Start();

                            //2016/穎驊註解，經由與恩正討論，現在無論有缺漏，全部人的資料都要顯示出來，故將條件註解掉
                            //if (cp.GetErrorCount() > 0)
                            //{
                            if (!cs.NonInputCompleteDict.ContainsKey(cpGName))
                                cs.NonInputCompleteDict.Add(cpGName, cp.GetMessage());
                            //}

                            cs.All_ErrorCount += cp.GetErrorCount();
                            cs.All_TotalCount += cp.GetTotalCount();

                        }
                    }
                }

                _bgLoadData.ReportProgress(80);

                // 讀取樣版
                Workbook wb = new Workbook(new MemoryStream(Properties.Resources.綜合紀錄表輸入進度表樣版));

                // 綜合紀錄索引
                Dictionary<string, int> gpColIdx = new Dictionary<string, int>();
                int col = 6;
                foreach (string cpName in cpGNameList)
                    gpColIdx.Add(cpName, col++);

                //2016/9/9 穎驊註解，此為給文華高中 輔導系統2.0 的欄位項目
                //個人資料	監護人資料	尊親屬資料	兄弟姊妹資料	身高及體重	家庭訊息	學習	幹部資訊	自我認識	生活感想	畢業後規劃	自傳	完成百分比


                #region 詳細資料清單
                List<string> chkItems1 = new List<string>();
                List<string> chkItems2 = new List<string>();
                List<string> chkItems3 = new List<string>();
                List<string> chkItems4 = new List<string>();
                List<string> chkItems5 = new List<string>();
                List<string> chkItems6 = new List<string>();
                List<string> chkItems7 = new List<string>();
                List<string> chkItems8 = new List<string>();
                List<string> chkItems9 = new List<string>();
                List<string> chkItems10 = new List<string>();
                List<string> chkItems11 = new List<string>();
                List<string> chkItems12 = new List<string>();

                //個人資料
                chkItems1.Add("手機號碼");
                chkItems1.Add("血型");
                chkItems1.Add("宗教");
                chkItems1.Add("生理缺陷");
                chkItems1.Add("曾患特殊疾病");
                chkItems1.Add("原住民血統");
                //監護人資料
                chkItems2.Add("監護人_姓名");
                chkItems2.Add("監護人_性別");
                chkItems2.Add("監護人_關係");
                chkItems2.Add("監護人_電話");
                chkItems2.Add("監護人_通訊地址");
                //尊親屬資料
                chkItems3.Add("title");
                chkItems3.Add("Name");
                chkItems3.Add("birth_year");
                chkItems3.Add("is_alive");
                chkItems3.Add("Phone");
                chkItems3.Add("Job");
                chkItems3.Add("Institute");
                chkItems3.Add("job_title");
                chkItems3.Add("edu_degree");
                chkItems3.Add("National");
                chkItems3.Add("cell_phone");
                //兄弟姊妹資料
                chkItems4.Add("title");
                chkItems4.Add("Name");
                chkItems4.Add("birth_year");
                chkItems4.Add("school_name");
                chkItems4.Add("remark");
                //身高及體重
                chkItems5.Add("s1a");
                chkItems5.Add("s1b");
                chkItems5.Add("s2a");
                chkItems5.Add("s2b");
                chkItems5.Add("s3a");
                chkItems5.Add("s3b");
                //家庭訊息
                chkItems6.Add("父母關係");
                chkItems6.Add("家庭氣氛");
                chkItems6.Add("父親管教方式");
                chkItems6.Add("母親管教方式");
                chkItems6.Add("居住環境");
                chkItems6.Add("本人住宿");
                chkItems6.Add("經濟狀況");
                chkItems6.Add("每星期零用錢");
                chkItems6.Add("我覺得是否足夠");
                //學習                
                chkItems7.Add("最喜歡的學科");
                chkItems7.Add("最感困難的學科");
                chkItems7.Add("特殊專長");
                chkItems7.Add("休閒興趣");
                chkItems7.Add("特殊專長_樂器演奏");
                chkItems7.Add("特殊專長_外語能力");
                //幹部資訊
                chkItems8.Add("班級幹部s1a");
                chkItems8.Add("社團幹部s1a");
                chkItems8.Add("班級幹部s1b");
                chkItems8.Add("社團幹部s1b");
                chkItems8.Add("班級幹部s2a");
                chkItems8.Add("社團幹部s2a");
                chkItems8.Add("班級幹部s2b");
                chkItems8.Add("社團幹部s2b");
                chkItems8.Add("班級幹部s3a");
                chkItems8.Add("社團幹部s3a");
                chkItems8.Add("班級幹部s3b");
                chkItems8.Add("社團幹部s3b");
                //自我認識			
                chkItems9.Add("個性");
                chkItems9.Add("需要改進的地方");
                chkItems9.Add("優點");
                chkItems9.Add("填寫日期");
                //生活感想
                chkItems10.Add("內容1");
                chkItems10.Add("內容2");
                chkItems10.Add("內容3");
                chkItems10.Add("填寫日期");
                //畢業後規劃
                chkItems11.Add("升學意願");
                chkItems11.Add("就業意願");
                chkItems11.Add("參加職業訓練");
                chkItems11.Add("受訓地區");
                chkItems11.Add("將來職業");
                chkItems11.Add("就業地區");
                //自傳
                chkItems12.Add("家中最了解我的人");
                chkItems12.Add("家中最了解我的人_因為");
                chkItems12.Add("我在家中最怕的人是");
                chkItems12.Add("我在家中最怕的人是_因為");
                chkItems12.Add("常指導我做功課的人");
                chkItems12.Add("讀過且印象最深刻的課外書");
                chkItems12.Add("喜歡的人");
                chkItems12.Add("喜歡的人_因為");
                chkItems12.Add("最要好的朋友");
                chkItems12.Add("他是怎樣的人");
                chkItems12.Add("最喜歡做的事");
                chkItems12.Add("最喜歡做的事_因為");
                chkItems12.Add("最不喜歡做的事");
                chkItems12.Add("最不喜歡做的事_因為");
                chkItems12.Add("國中時的學校生活");
                chkItems12.Add("最快樂的回憶");
                chkItems12.Add("最痛苦的回憶");
                chkItems12.Add("最足以描述自己的幾句話");
                chkItems12.Add("我覺得我的優點是");
                chkItems12.Add("我覺得我的缺點是");
                chkItems12.Add("最喜歡的國小（國中）老師");
                chkItems12.Add("最喜歡的國小（國中）老師__因為");
                chkItems12.Add("小學（國中）老師或同學常說我是");
                chkItems12.Add("小學（國中）時我曾在班上登任過的職務有");
                chkItems12.Add("我在小學（國中）得過的獎有");
                chkItems12.Add("我覺得我自己的過去最滿意的是");
                chkItems12.Add("我排遣休閒時間的方法是");
                chkItems12.Add("我最難忘的一件事是");
                chkItems12.Add("自傳");
                chkItems12.Add("自我的心聲_一年級_我目前遇到最大的困難是");
                chkItems12.Add("自我的心聲_一年級_我目前最需要的協助是");
                chkItems12.Add("自我的心聲_二年級_我目前遇到最大的困難是");
                chkItems12.Add("自我的心聲_二年級_我目前最需要的協助是");
                chkItems12.Add("自我的心聲_三年級_我目前遇到最大的困難是");
                chkItems12.Add("自我的心聲_三年級_我目前最需要的協助是"); 
                #endregion
             
                int rowIdx = 1;

                //是否唯獨子
                bool I_am_the_onle_child = false;

                foreach (ClassStudent cs in ClassStudents)
                {
                    int Coldx = 6;

                    //2016/穎驊註解，經由與恩正討論，現在無論有缺漏，全部人的資料都要顯示出來，故將條件註解掉
                    //// 有缺才填入
                    //if (cs.NonInputCompleteDict.Count > 0)
                    //{
                    // 填入缺漏資料
                    foreach (string key in cs.NonInputCompleteDict.Keys)
                    {
                        if (gpColIdx.ContainsKey(key))
                            wb.Worksheets[0].Cells[rowIdx, gpColIdx[key]].PutValue(cs.NonInputCompleteDict[key]);
                    }


                    wb.Worksheets[0].Cells[rowIdx, 0].PutValue(cs.StudentNumber);
                    wb.Worksheets[0].Cells[rowIdx, 1].PutValue(cs.GradeYearDisplay);
                    wb.Worksheets[0].Cells[rowIdx, 2].PutValue(cs.ClassName);
                    wb.Worksheets[0].Cells[rowIdx, 3].PutValue(cs.SeatNo);
                    wb.Worksheets[0].Cells[rowIdx, 4].PutValue(cs.StudentName);
                    //2016/9/9 穎驊新增，計算學生完成百分比
                    decimal inputCompletePrecent = Math.Round((((decimal)cs.All_TotalCount - (decimal)(cs.All_ErrorCount)) / ((decimal)cs.All_TotalCount)), 2, MidpointRounding.AwayFromZero);
                    wb.Worksheets[0].Cells[rowIdx, 5].PutValue(inputCompletePrecent);



                    //2016/9/20，穎驊新增，顯示學生填答詳細資料在第二張Sheet
                    wb.Worksheets[1].Cells[rowIdx, 0].PutValue(cs.StudentNumber);
                    wb.Worksheets[1].Cells[rowIdx, 1].PutValue(cs.GradeYearDisplay);
                    wb.Worksheets[1].Cells[rowIdx, 2].PutValue(cs.ClassName);
                    wb.Worksheets[1].Cells[rowIdx, 3].PutValue(cs.SeatNo);
                    wb.Worksheets[1].Cells[rowIdx, 4].PutValue(cs.StudentName);

                    //2016/9/23 穎驊新增，將在詳細資料Sheet 內，加入公式，動態計算完成百分比。  
                    //可以依每行數增加而動態改變組成字串
                    wb.Worksheets[1].Cells[rowIdx, 5].Formula = "=COUNTA($G" + (rowIdx+1) + ":$XFD"+(rowIdx+1)+")/COUNTA($G$1:$XFD$1)";
                    

                    #region 自Gobal的資料整理出每一個學生的檔案塞給StuData
                    if (Gobal._single_recordDict.ContainsKey(cs.StudentID))
                    {
                        foreach (DataRow dr in Gobal._single_recordDict[cs.StudentID])
                        {
                            if (!StuData.ContainsKey(cs.StudentID))
                            {
                                StuData.Add(cs.StudentID, new Dictionary<string, string>());

                                if (!StuData[cs.StudentID].ContainsKey(dr["key"].ToString())) 
                                {
                                    StuData[cs.StudentID].Add(dr["key"].ToString(), dr["data"].ToString().Trim());
                                }
                                
                            }
                            else
                            {
                                if (!StuData[cs.StudentID].ContainsKey(dr["key"].ToString()))
                                {
                                    StuData[cs.StudentID].Add(dr["key"].ToString(), dr["data"].ToString().Trim());
                                }
                            }
                        }
                    }

                    if (Gobal._multiple_recordDict.ContainsKey(cs.StudentID))
                    {
                        foreach (DataRow dr in Gobal._multiple_recordDict[cs.StudentID])
                        {
                            if (!StuData.ContainsKey(cs.StudentID))
                            {
                                StuData.Add(cs.StudentID, new Dictionary<string, string>());

                                StuData[cs.StudentID].Add(dr["key"].ToString(), dr["data"].ToString().Trim());
                            }
                            else
                            {
                                if (StuData[cs.StudentID].ContainsKey(dr["key"].ToString()) && dr["data"].ToString().Trim()!="")
                                {
                                    StuData[cs.StudentID][dr["key"].ToString()] += "、"+dr["data"].ToString().Trim();                                
                                }
                                else
                                {
                                StuData[cs.StudentID].Add(dr["key"].ToString(), dr["data"].ToString().Trim());                                
                                }                                
                            }
                        }
                    }

                    if (Gobal._priority_dataDict.ContainsKey(cs.StudentID))
                    {
                        foreach (DataRow dr in Gobal._priority_dataDict[cs.StudentID])
                        {
                            if (!StuData.ContainsKey(cs.StudentID))
                            {
                                StuData.Add(cs.StudentID, new Dictionary<string, string>());

                                if (dr["p1"].ToString().Trim() != "" && !StuData[cs.StudentID].ContainsKey(dr["key"].ToString()))
                                {
                                    StuData[cs.StudentID].Add(dr["key"].ToString(), dr["p1"].ToString().Trim() + "、" + dr["p2"].ToString().Trim() + "、" + dr["p3"].ToString().Trim());
                                }

                            }
                            else
                            {
                                if (dr["p1"].ToString().Trim() != "" && !StuData[cs.StudentID].ContainsKey(dr["key"].ToString()))
                                {
                                    StuData[cs.StudentID].Add(dr["key"].ToString(), dr["p1"].ToString().Trim() + "、" + dr["p2"].ToString().Trim() + "、" + dr["p3"].ToString().Trim());
                                }
                            }
                        }
                    }
              

                    if (Gobal._relativeDict.ContainsKey(cs.StudentID))
                    {
                        foreach (DataRow dr in Gobal._relativeDict[cs.StudentID])
                        {
                            if (!StuData.ContainsKey(cs.StudentID))
                            {
                                StuData.Add(cs.StudentID, new Dictionary<string, string>());

                                foreach (string ssKey in chkItems3)
                                {

                                    if (ssKey == "is_alive" && dr[ssKey].ToString().Trim() == "true")
                                    {
                                        StuData[cs.StudentID].Add("尊親屬資料_" + ssKey, "存");                                
                                    }
                                    else if (ssKey == "is_alive" && dr[ssKey].ToString().Trim() == "false")
                                    {
                                        StuData[cs.StudentID].Add("尊親屬資料_" + ssKey, "歿");
                                    }
                                    else 
                                    {
                                        StuData[cs.StudentID].Add("尊親屬資料_" + ssKey, "\n"+dr[ssKey].ToString().Trim());                                
                                    }
                                    
                                }                                
                            }
                            else
                            {
                                foreach (string ssKey in chkItems3)
                                {
                                    if (!StuData[cs.StudentID].ContainsKey("尊親屬資料_" + ssKey)) 
                                    {
                                        if (ssKey == "is_alive" && dr[ssKey].ToString().Trim() == "true")
                                        {
                                            StuData[cs.StudentID].Add("尊親屬資料_" + ssKey, "存");
                                        }
                                        else if (ssKey == "is_alive" && dr[ssKey].ToString().Trim() == "false")
                                        {
                                            StuData[cs.StudentID].Add("尊親屬資料_" + ssKey, "歿");
                                        }
                                        else 
                                        {
                                            StuData[cs.StudentID].Add("尊親屬資料_" + ssKey, "\n"+dr[ssKey].ToString().Trim());
                                        }
                                        continue;
                                    }
                                    if (StuData[cs.StudentID].ContainsKey("尊親屬資料_" + ssKey) )
                                    {
                                        if (ssKey == "is_alive" && dr[ssKey].ToString().Trim() == "true")
                                        {
                                            StuData[cs.StudentID]["尊親屬資料_" + ssKey] += "\n" +"存";
                                        }
                                        else if (ssKey == "is_alive" && dr[ssKey].ToString().Trim() == "false")
                                        {
                                            StuData[cs.StudentID]["尊親屬資料_" + ssKey] += "\n" + "歿";
                                        }
                                        else
                                        {
                                            StuData[cs.StudentID]["尊親屬資料_" + ssKey] += "\n" + dr[ssKey].ToString().Trim();
                                        }
                                        
                                    }
                                    
                                }
                            }


                        }
                    }
                    if (Gobal._semester_dataDict.ContainsKey(cs.StudentID))
                    {
                        foreach (DataRow dr in Gobal._semester_dataDict[cs.StudentID])
                        {
                            if (!StuData.ContainsKey(cs.StudentID))
                            {
                                StuData.Add(cs.StudentID, new Dictionary<string, string>());

                                foreach (string ssKey in chkItems5) 
                                {
                                    if (dr[ssKey] != null) 
                                    {
                                        if (!StuData[cs.StudentID].ContainsKey(dr["key"].ToString() + ssKey))
                                        {
                                            StuData[cs.StudentID].Add(dr["key"].ToString() + ssKey, dr[ssKey].ToString().Trim());
                                        }
                                    }                                    
                                }                                
                            }
                            else
                            {
                                foreach (string ssKey in chkItems5)
                                {
                                    if (dr[ssKey] != null && dr[ssKey] + "" != "" && !StuData[cs.StudentID].ContainsKey(dr["key"].ToString() + ssKey))
                                    {
                                        StuData[cs.StudentID].Add(dr["key"].ToString()+ssKey, dr[ssKey].ToString().Trim());
                                    }
                                }
                            }
                        }
                    }

                    if (Gobal._siblingDict.ContainsKey(cs.StudentID))
                    {

                        foreach (DataRow dr in Gobal._siblingDict[cs.StudentID])
                        {
                            if (!StuData.ContainsKey(cs.StudentID))
                            {
                                StuData.Add(cs.StudentID, new Dictionary<string, string>());

                                foreach (string ssKey in chkItems4) 
                                {
                                    StuData[cs.StudentID].Add("兄弟姊妹資料_"+ssKey, dr[ssKey].ToString().Trim());
                                
                                }
                                
                            }
                            else
                            {
                                foreach (string ssKey in chkItems4)
                                {
                                    if (!StuData[cs.StudentID].ContainsKey("兄弟姊妹資料_" + ssKey))
                                    {
                                        StuData[cs.StudentID].Add("兄弟姊妹資料_" + ssKey, dr[ssKey].ToString().Trim());
                                    }
                                    if (StuData[cs.StudentID].ContainsKey("兄弟姊妹資料_" + ssKey) && StuData[cs.StudentID]["兄弟姊妹資料_" + ssKey] != dr[ssKey].ToString().Trim())
                                    {
                                        StuData[cs.StudentID]["兄弟姊妹資料_" + ssKey] += "\n"+dr[ssKey].ToString().Trim();
                                    }
                                    
                                }
                            }
                        }
                    }
                    
                    if (Gobal._yearly_dataDict.ContainsKey(cs.StudentID))
                    {
                        foreach (DataRow dr in Gobal._yearly_dataDict[cs.StudentID])
                        {
                            if (!StuData.ContainsKey(cs.StudentID))
                            {
                                StuData.Add(cs.StudentID, new Dictionary<string, string>());

                                if (!StuData[cs.StudentID].ContainsKey(dr["key"].ToString())) 
                                {
                                    StuData[cs.StudentID].Add(dr["key"].ToString(), dr["g1"].ToString().Trim());
                                }                                
                            }
                            else
                            {
                                if (!StuData[cs.StudentID].ContainsKey(dr["key"].ToString()))
                                {
                                    StuData[cs.StudentID].Add(dr["key"].ToString(), dr["g1"].ToString().Trim());
                                }
                            }
                        }
                    } 
                    #endregion


                    //將沒有資料的欄位填成粉紅色，使使用者比較容易使用
                    //另外一提style1.Pattern一定要設定，否則會填色失敗。
                    Style style1 = wb.Styles[wb.Styles.Add()];
                    StyleFlag flag = new StyleFlag();

                    flag.All = true;
                    style1.ForegroundColor = Color.LightPink;
                    style1.BackgroundColor = Color.LightPink;
                    style1.Pattern = BackgroundType.HorizontalStripe;
                        

                    //2016/9/26 穎驊註解，下面開始針對Excel 進行填值，假如沒有資料 不會再另外填""空字串，直接填色，
                    //因為就本例而言，填空字串會造成公式"=COUNTA($G" + (rowIdx+1) + ":$XFD"+(rowIdx+1)+")/COUNTA($G$1:$XFD$1)"; 的誤算。
                    #region 開始填詳細資料
                    if (StuData.ContainsKey(cs.StudentID))
                    {
                        //個人資料
                        foreach (String ssKey in chkItems1)
                        {
                            if (StuData[cs.StudentID].ContainsKey("本人概況_" + ssKey))
                            {                                
                                if (StuData[cs.StudentID]["本人概況_" + ssKey] == "")
                                {
                                    wb.Worksheets[1].Cells[rowIdx, Coldx].SetStyle(style1);
                                }
                                else 
                                {
                                    wb.Worksheets[1].Cells[rowIdx, Coldx].PutValue(StuData[cs.StudentID]["本人概況_" + ssKey]);
                                }
                            }
                            else
                            {
                                wb.Worksheets[1].Cells[rowIdx, Coldx].SetStyle(style1);
                            }

                            Coldx++;

                        }

                        //監護人資料
                        foreach (String ssKey in chkItems2)
                        {
                            if (StuData[cs.StudentID].ContainsKey("家庭狀況_" + ssKey))
                            {
                                if (StuData[cs.StudentID]["家庭狀況_" + ssKey] == "")
                                {
                                    wb.Worksheets[1].Cells[rowIdx, Coldx].SetStyle(style1, flag);
                                }
                                else 
                                {
                                    wb.Worksheets[1].Cells[rowIdx, Coldx].PutValue(StuData[cs.StudentID]["家庭狀況_" + ssKey]);
                                }                                
                            }
                            else
                            {                               
                                wb.Worksheets[1].Cells[rowIdx, Coldx].SetStyle(style1, flag);
                            }

                            Coldx++;

                        }

                        //尊親屬資料

                        foreach (String ssKey in chkItems3)
                        {
                            if (StuData[cs.StudentID].ContainsKey("尊親屬資料_" + ssKey))
                            {
                                if (StuData[cs.StudentID]["尊親屬資料_" + ssKey] == "")
                                {
                                    wb.Worksheets[1].Cells[rowIdx, Coldx].SetStyle(style1, flag);
                                }
                                else 
                                {
                                    wb.Worksheets[1].Cells[rowIdx, Coldx].PutValue(StuData[cs.StudentID]["尊親屬資料_" + ssKey]);
                                }                                
                            }
                            else
                            {                        
                                wb.Worksheets[1].Cells[rowIdx, Coldx].SetStyle(style1, flag);
                            }

                            Coldx++;

                        }

                        if (StuData[cs.StudentID].ContainsKey("家庭狀況_兄弟姊妹_排行"))
                        {
                            if (StuData[cs.StudentID]["家庭狀況_兄弟姊妹_排行"] == "")
                            {
                                wb.Worksheets[1].Cells[rowIdx, Coldx].PutValue("我是獨子女");

                                I_am_the_onle_child = true;
                            }
                            else 
                            {
                                wb.Worksheets[1].Cells[rowIdx, Coldx].PutValue(StuData[cs.StudentID]["家庭狀況_兄弟姊妹_排行"]);

                                I_am_the_onle_child = false;
                            }

                            Coldx++;
                        }
                        else
                        {
                            wb.Worksheets[1].Cells[rowIdx, Coldx].SetStyle(style1, flag);

                            Coldx++;
                        
                        }


                        if (!I_am_the_onle_child)
                        {
                            //兄弟姊妹資料
                            foreach (String ssKey in chkItems4)
                            {
                                if (StuData[cs.StudentID].ContainsKey("兄弟姊妹資料_" + ssKey))
                                {
                                    if (StuData[cs.StudentID]["兄弟姊妹資料_" + ssKey] == "")
                                    {
                                        wb.Worksheets[1].Cells[rowIdx, Coldx].SetStyle(style1, flag);
                                    }
                                    else
                                    {
                                        wb.Worksheets[1].Cells[rowIdx, Coldx].PutValue(StuData[cs.StudentID]["兄弟姊妹資料_" + ssKey]);
                                    }
                                }
                                else
                                {
                                    wb.Worksheets[1].Cells[rowIdx, Coldx].SetStyle(style1, flag);
                                }

                                Coldx++;

                            }
                        }
                        else 
                        {
                            foreach (String ssKey in chkItems4)
                            {
                                wb.Worksheets[1].Cells[rowIdx, Coldx].PutValue("無");

                                Coldx++;
                            }
                        }

                        //身高及體重

                        foreach (String ssKey in chkItems5)
                        {
                            if (StuData[cs.StudentID].ContainsKey("本人概況_身高" + ssKey))
                            {
                                if (StuData[cs.StudentID]["本人概況_身高" + ssKey] == "")
                                {
                                    wb.Worksheets[1].Cells[rowIdx, Coldx].SetStyle(style1, flag);
                                }
                                else 
                                {
                                    wb.Worksheets[1].Cells[rowIdx, Coldx].PutValue(StuData[cs.StudentID]["本人概況_身高" + ssKey]);
                                }                                
                            }
                            else
                            {                   
                                wb.Worksheets[1].Cells[rowIdx, Coldx].SetStyle(style1, flag);
                            }

                            if (StuData[cs.StudentID].ContainsKey("本人概況_體重" + ssKey))
                            {
                                if (StuData[cs.StudentID]["本人概況_體重" + ssKey] == "")
                                {
                                    wb.Worksheets[1].Cells[rowIdx, Coldx + 1].SetStyle(style1, flag);
                                }
                                else 
                                {
                                    wb.Worksheets[1].Cells[rowIdx, Coldx + 1].PutValue(StuData[cs.StudentID]["本人概況_體重" + ssKey]);
                                }                                
                            }
                            else
                            {                    
                                wb.Worksheets[1].Cells[rowIdx, Coldx+1].SetStyle(style1, flag);
                            }

                            Coldx += 2;

                        }

                        // 家庭訊息                        
                        foreach (String ssKey in chkItems6)
                        {
                            if (StuData[cs.StudentID].ContainsKey("家庭狀況_" + ssKey))
                            {
                                if (StuData[cs.StudentID]["家庭狀況_" + ssKey] == "")
                                {
                                    wb.Worksheets[1].Cells[rowIdx, Coldx].SetStyle(style1, flag);
                                }
                                else 
                                {
                                    wb.Worksheets[1].Cells[rowIdx, Coldx].PutValue(StuData[cs.StudentID]["家庭狀況_" + ssKey]);
                                }
                                
                            }
                            else
                            {                        
                                wb.Worksheets[1].Cells[rowIdx, Coldx].SetStyle(style1, flag);
                            }

                            Coldx++;

                        }

                        // 學習                        
                        foreach (String ssKey in chkItems7)
                        {
                            if (StuData[cs.StudentID].ContainsKey("學習狀況_" + ssKey))
                            {
                                if (StuData[cs.StudentID]["學習狀況_" + ssKey] == "")
                                {
                                    wb.Worksheets[1].Cells[rowIdx, Coldx].SetStyle(style1, flag);
                                }
                                else 
                                {
                                    wb.Worksheets[1].Cells[rowIdx, Coldx].PutValue(StuData[cs.StudentID]["學習狀況_" + ssKey]);
                                }
                                
                            }
                            else
                            {                              
                                wb.Worksheets[1].Cells[rowIdx, Coldx].SetStyle(style1, flag);
                            }

                            Coldx++;

                        }

                        // 幹部資訊                        
                        foreach (String ssKey in chkItems8)
                        {
                            if (StuData[cs.StudentID].ContainsKey("學習狀況_" + ssKey))
                            {
                                if (StuData[cs.StudentID]["學習狀況_" + ssKey] == "")
                                {
                                    wb.Worksheets[1].Cells[rowIdx, Coldx].SetStyle(style1, flag);
                                }
                                else
                                {
                                    wb.Worksheets[1].Cells[rowIdx, Coldx].PutValue(StuData[cs.StudentID]["學習狀況_" + ssKey]);
                                }                                
                            }
                            else
                            {                        
                                wb.Worksheets[1].Cells[rowIdx, Coldx].SetStyle(style1, flag);
                            }

                            Coldx++;

                        }

                        //自我認識
                        foreach (String ssKey in chkItems9)
                        {
                            if (StuData[cs.StudentID].ContainsKey("自我認識_" + ssKey + "_" + cs.GradeYear))
                            {
                                if (StuData[cs.StudentID]["自我認識_" + ssKey + "_" + cs.GradeYear] == "")
                                {
                                    wb.Worksheets[1].Cells[rowIdx, Coldx].SetStyle(style1, flag);
                                }
                                else 
                                {
                                    wb.Worksheets[1].Cells[rowIdx, Coldx].PutValue(StuData[cs.StudentID]["自我認識_" + ssKey + "_" + cs.GradeYear]);
                                }                                
                            }
                            else
                            {               
                                wb.Worksheets[1].Cells[rowIdx, Coldx].SetStyle(style1, flag);
                            }

                            Coldx++;

                        }

                        //生活感想
                        foreach (String ssKey in chkItems10)
                        {
                            if (StuData[cs.StudentID].ContainsKey("生活感想_" + ssKey + "_" + cs.GradeYear))
                            {
                                if (StuData[cs.StudentID]["生活感想_" + ssKey + "_" + cs.GradeYear] == "")
                                {
                                    wb.Worksheets[1].Cells[rowIdx, Coldx].SetStyle(style1, flag);
                                }
                                else 
                                {
                                    wb.Worksheets[1].Cells[rowIdx, Coldx].PutValue(StuData[cs.StudentID]["生活感想_" + ssKey + "_" + cs.GradeYear]);
                                }                                
                            }
                            else
                            {                          
                                wb.Worksheets[1].Cells[rowIdx, Coldx].SetStyle(style1, flag);
                            }

                            Coldx++;

                        }

                        //畢業後規劃
                        foreach (String ssKey in chkItems11)
                        {
                            if (StuData[cs.StudentID].ContainsKey("畢業後計畫_" + ssKey))
                            {
                                if (StuData[cs.StudentID]["畢業後計畫_" + ssKey] == "")
                                {
                                    wb.Worksheets[1].Cells[rowIdx, Coldx].SetStyle(style1, flag);
                                }
                                else 
                                {
                                    wb.Worksheets[1].Cells[rowIdx, Coldx].PutValue(StuData[cs.StudentID]["畢業後計畫_" + ssKey]);
                                }                                
                            }
                            else
                            {        
                                wb.Worksheets[1].Cells[rowIdx, Coldx].SetStyle(style1, flag);
                            }

                            Coldx++;

                        }

                        //自傳
                        foreach (String ssKey in chkItems12)
                        {
                            if (StuData[cs.StudentID].ContainsKey("自傳_" + ssKey))
                            {                                
                                if (StuData[cs.StudentID]["自傳_" + ssKey] == "")
                                {
                                    wb.Worksheets[1].Cells[rowIdx, Coldx].SetStyle(style1, flag);
                                }
                                else 
                                {
                                    wb.Worksheets[1].Cells[rowIdx, Coldx].PutValue(StuData[cs.StudentID]["自傳_" + ssKey]);
                                }
                            }
                       
                            else
                            {                
                                wb.Worksheets[1].Cells[rowIdx, Coldx].SetStyle(style1, flag);
                            }

                            Coldx++;

                        }
                    }
                    else 
                    {
                        int i_for_total_imcomplete = 0;

                        i_for_total_imcomplete += chkItems1.Count;
                        i_for_total_imcomplete += chkItems2.Count;
                        i_for_total_imcomplete += chkItems3.Count;
                        i_for_total_imcomplete += chkItems4.Count;
                        //身高體重有兩次
                        i_for_total_imcomplete += chkItems5.Count*2;
                        i_for_total_imcomplete += chkItems6.Count;
                        i_for_total_imcomplete += chkItems7.Count;
                        i_for_total_imcomplete += chkItems8.Count;
                        i_for_total_imcomplete += chkItems9.Count;
                        i_for_total_imcomplete += chkItems10.Count;
                        i_for_total_imcomplete += chkItems11.Count;
                        i_for_total_imcomplete += chkItems12.Count;

                        for (int i = 0; i < i_for_total_imcomplete; i++) 
                        {
                            wb.Worksheets[1].Cells[rowIdx, Coldx+i].SetStyle(style1, flag);
                        
                        }
                                                                    
                    }
 

                    #endregion

                    rowIdx++;
                    //}

                }
                _bgLoadData.ReportProgress(95);

                wb.Worksheets[0].AutoFitColumns();


                //2016/9/26 穎驊註解，因為有更好用的第二張Sheet 出來可供參照，但與第一頁的完成%數 會因為一些狀況而有誤差，為了避免使用者誤解，將第一張表刪去
                wb.Worksheets.RemoveAt(0);
                
                

                ////2016/9/21 穎驊筆記，處理使Excel能自動換行

                //Style style1 = wb.Styles[wb.Styles.Add()];
                //StyleFlag flag = new StyleFlag();

                //flag.All = true;
                //style1.IsTextWrapped = true;

                //wb.Worksheets[1].Cells.ApplyStyle(style1, flag);


                e.Result = wb;

                _bgLoadData.ReportProgress(100);
            }
            catch (Exception ex)
            {


                e.Cancel = true;
            }
        }

    }
}
