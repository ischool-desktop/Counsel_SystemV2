using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

using FISCA.Presentation.Controls;
using Aspose.Cells;
using K12.Data;
using FISCA.DSAClient;
using FISCA.DSAUtil;
using FISCA.UDT;


namespace Psychological_Test_Import_whsh.Forms
{
    public partial class ImportStudent_MajorInTest : BaseForm
    {

        BackgroundWorker _bw = new BackgroundWorker();

        public string source_data = "";
        string target_grade_year = "";

        bool useIDNumberCheck = false;

        DateTime dt;


        public ImportStudent_MajorInTest()
        {
            InitializeComponent();
        }
                                
        // 選擇匯入檔案
        private void buttonX1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ope = new OpenFileDialog();
            ope.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            if (ope.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }
            else
            {
                textBoxX1.Text = ope.FileName;                
            }
        }

        // 匯入
        private void buttonX2_Click(object sender, EventArgs e)
        {
            source_data = textBoxX1.Text;

            useIDNumberCheck = checkBox1.Checked;

            dt = dateTimeInput1.Value;

            // 若沒選取來源檔案，中止程序
            if (source_data == "")
            {
                MsgBox.Show("請選擇來源檔案", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 若沒選取送出日期，中止程序
            if ("" + dateTimeInput1.Text == "")
            {
                MsgBox.Show("請選擇送出日期", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }


            // 暫停UI 功能
            textBoxX1.Enabled = false;
                        
            buttonX1.Enabled = false;
            buttonX2.Enabled = false;
            
            //加入背景執行序
            _bw.DoWork += new DoWorkEventHandler(_bkWork_DoWork);
            _bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_worker_RunWorkerCompleted);
            _bw.ProgressChanged += new ProgressChangedEventHandler(_worker_ProgressChanged);
            _bw.WorkerReportsProgress = true;

            _bw.RunWorkerAsync();
                        
        }

        private void _bkWork_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                // 0.載入匯入檔案
                Workbook wb = new Workbook(source_data);

                //2017/3/8 穎驊筆記，大考中心提供的興趣測驗 EXCEL 匯入檔 格式過於老舊，
                //雖然其副檔名 為xls， 但其代表的其實是舊檔格式MicroSoft Excel5.0/95，
                //而非我們所了解的 MicroSoft Excel 97~2003，
                //而我們公司所用的處理Excel 、Word 的Aspose API
                //完全不支援MicroSoft Excel5.0/95的編碼格式，
                //必須要請使用者 另外再把檔案存成MicroSoft Excel 97~2003 才可以匯入，
                //這部分需要在介面上提醒，以及加強與對方的教學。


                // 文件打開舊方法，已過時， 現在統一使用上面的new Workbook(textBoxX1.Text) 的方式，另外 請使用最新 Aspose.Cell_201402 ，可以避免很多讀取錯誤Bug
                //wb.Open(textBoxX1.Text, FileFormatType.Excel2007Xlsx);

                Worksheet ws = wb.Worksheets[0];

                Cells cells = ws.Cells;

                #region 驗證欄位 List
                List<string> verifiedDomainList1 = new List<string>();

                verifiedDomainList1.Add("學校班級名稱");
                verifiedDomainList1.Add("學生姓名");
                verifiedDomainList1.Add("學生座號");
                verifiedDomainList1.Add("身分證字號");
                verifiedDomainList1.Add("1");
                verifiedDomainList1.Add("2");
                verifiedDomainList1.Add("3");
                verifiedDomainList1.Add("4");
                verifiedDomainList1.Add("5");
                verifiedDomainList1.Add("6");
                verifiedDomainList1.Add("7");
                verifiedDomainList1.Add("8");
                verifiedDomainList1.Add("1 數學分數");
                verifiedDomainList1.Add("2 物理分數");
                verifiedDomainList1.Add("3 化學分數");
                verifiedDomainList1.Add("4 資訊電子分數");
                verifiedDomainList1.Add("5 通訊電信分數");
                verifiedDomainList1.Add("6 工程科技分數");                
                verifiedDomainList1.Add("7 機械分數");
                verifiedDomainList1.Add("8 建築營造分數");
                verifiedDomainList1.Add("9 設計分數");
                verifiedDomainList1.Add("10 生命科學分數");
                verifiedDomainList1.Add("11 醫學分數");
                verifiedDomainList1.Add("12 生資食科分數");
                verifiedDomainList1.Add("13 地球環境分數");
                verifiedDomainList1.Add("14 藝術分數");
                verifiedDomainList1.Add("15 歷史文化分數");
                verifiedDomainList1.Add("16 傳播媒體分數");
                verifiedDomainList1.Add("17 教育訓練分數");
                verifiedDomainList1.Add("18 心理學分數");
                verifiedDomainList1.Add("19 社會人類分數");
                verifiedDomainList1.Add("20 哲學宗教分數");
                verifiedDomainList1.Add("21 治療諮商分數");
                verifiedDomainList1.Add("22 語文文學分數");
                verifiedDomainList1.Add("23 外國語文分數");
                verifiedDomainList1.Add("24 人力資源分數");
                verifiedDomainList1.Add("25 顧客服務分數");
                verifiedDomainList1.Add("26 管理分數");
                verifiedDomainList1.Add("27 銷售行銷分數");
                verifiedDomainList1.Add("28 經濟會計分數");
                verifiedDomainList1.Add("29 法律政治分數");
                verifiedDomainList1.Add("30 行政分數");


                #endregion
                                

                // 錯誤資料List
                List<string> errorList = new List<string>();


                // 1.驗證資料

                // 建立身分字號  對應 StudentID 對照表
                List<StudentRecord> allStudentList = K12.Data.Student.SelectAll();

                //<身分字號，學生系統ID>
                Dictionary<string, string> ID_to_refstudentID_Dict = new Dictionary<string, string>();

                foreach (StudentRecord sr in allStudentList)
                {
                    if (!ID_to_refstudentID_Dict.ContainsKey(sr.IDNumber))
                    {
                        ID_to_refstudentID_Dict.Add(sr.IDNumber, sr.ID);

                    }
                }


                // 1.1 驗證欄位標題

                #region 驗證欄位標題
                int i = 0;

                foreach (string domain in verifiedDomainList1)
                {
                    if ("" + cells[0,i].Value != verifiedDomainList1[i])
                    {
                        errorList.Add("選擇檔案EXCEL 欄位於 第一列 第" + (i+1)  + "行處，與匯入格式不符，請確認");
                    }              
                    i++;
                }

                i = 0;
                
                #endregion


                //1.2  驗證班級學生資料

                #region 驗證班級學生資料

                Dictionary<string, List<string>> studentDict_By_Class = new Dictionary<string, List<string>>();

                foreach (var row in ws.Cells.Rows)
                {
                    // 身分證字號 
                    if (useIDNumberCheck && row.Index > 0 && "" + cells[row.Index, 3].Value != "" && !row.IsBlank)
                    {
                        if (!ID_to_refstudentID_Dict.ContainsKey("" + cells[row.Index, 3].Value))
                        {
                            errorList.Add("第" + (row.Index + 1) + "列 第" + 4 + "行處，系統不存在此身分字號學生，請檢察。");
                        }                              
                    }

                    if (useIDNumberCheck && row.Index > 0 && "" + cells[row.Index, 3].Value == "" && !row.IsBlank)
                    {
                        errorList.Add("第" + (row.Index + 1) + "列 第" + 4 + "行處，該學生沒有身分字號，請檢察。");
                    }

                }
                #endregion


                //2.假如驗證不過，顯示訊息後，中止
                if (errorList.Count > 0)
                {
                    StringBuilder errorMessages = new StringBuilder(); ;

                    foreach (string errormessage in errorList)
                    {
                        errorMessages.AppendLine(errormessage);
                    }
                    MsgBox.Show(errorMessages.ToString(), "錯誤訊息", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);

                    return;
                }

                //3.讀取資料，寫入物件
                List<DAO.UDT_MajorIn_Test_Data_Def> interest_Test_Data_List = new List<DAO.UDT_MajorIn_Test_Data_Def>();

                // 警告資料List
                List<string> warningList = new List<string>();

                FISCA.UDT.AccessHelper accesshelper = new AccessHelper();

                int success_count = 0;

                foreach (var row in ws.Cells.Rows)
                {
                    string key = "" + cells[row.Index, 3].Value; // 身分字號
                    string studentID = "";

                    if (row.Index > 0 && !row.IsBlank)
                    {

                        if (ID_to_refstudentID_Dict.ContainsKey(key))
                        {
                            // 學生ID
                            studentID = "" + int.Parse(ID_to_refstudentID_Dict[key]);
                        }
                        else
                        {
                            warningList.Add("第" + (row.Index + 1) + "行，該身分字號的學生並不存在於本系統中，請檢察");
                        }
                        List<DAO.UDT_MajorIn_Test_Data_Def> dataList = accesshelper.Select<DAO.UDT_MajorIn_Test_Data_Def>("ref_student_id =" + "'" + studentID + "'");

                        if (dataList.Count == 0)
                        {
                            DAO.UDT_MajorIn_Test_Data_Def data = new DAO.UDT_MajorIn_Test_Data_Def();

                            #region 填值

                            //志願一
                            data.wish_1 = ("" + cells[row.Index, 4].Value == "" ? "_" : "" + cells[row.Index, 4].Value);
                            //志願二
                            data.wish_2 = ("" + cells[row.Index, 5].Value == "" ? "_" : "" + cells[row.Index, 5].Value);
                            //志願三
                            data.wish_3 = ("" + cells[row.Index, 6].Value == "" ? "_" : "" + cells[row.Index, 6].Value);
                            //志願四
                            data.wish_4 = ("" + cells[row.Index, 7].Value == "" ? "_" : "" + cells[row.Index, 7].Value);
                            //志願五
                            data.wish_5 = ("" + cells[row.Index, 8].Value == "" ? "_" : "" + cells[row.Index, 8].Value);
                            //志願六
                            data.wish_6 = ("" + cells[row.Index, 9].Value == "" ? "_" : "" + cells[row.Index, 9].Value);
                            //志願七
                            data.wish_7 = ("" + cells[row.Index, 10].Value == "" ? "_" : "" + cells[row.Index, 10].Value);
                            //志願八
                            data.wish_8 = ("" + cells[row.Index, 11].Value == "" ? "_" : "" + cells[row.Index, 11].Value);
                            //數學分數
                            data.math_score = ("" + cells[row.Index, 12].Value == "" ? "_" : "" + cells[row.Index, 12].Value);
                            //物理分數
                            data.phycics_score = ("" + cells[row.Index, 13].Value == "" ? "_" : "" + cells[row.Index, 13].Value);
                            //化學分數
                            data.chemistry_score = ("" + cells[row.Index, 14].Value == "" ? "_" : "" + cells[row.Index, 14].Value);
                            //資訊電子分數
                            data.information_electronics_score = ("" + cells[row.Index, 15].Value == "" ? "_" : "" + cells[row.Index, 15].Value);
                            //通訊電信分數
                            data.communication_telecommunications = ("" + cells[row.Index, 16].Value == "" ? "_" : "" + cells[row.Index, 16].Value);
                            //工程科技分數
                            data.engineer_technology_score = ("" + cells[row.Index, 17].Value == "" ? "_" : "" + cells[row.Index, 17].Value);
                            //機械分數
                            data.mechanism_score = ("" + cells[row.Index, 18].Value == "" ? "_" : "" + cells[row.Index, 18].Value);
                            //建築營造分數
                            data.building_construction_score = ("" + cells[row.Index, 19].Value == "" ? "_" : "" + cells[row.Index, 19].Value);
                            //設計分數
                            data.design_score = ("" + cells[row.Index, 20].Value == "" ? "_" : "" + cells[row.Index, 20].Value);
                            //生命科學分數
                            data.life_science_score = ("" + cells[row.Index, 21].Value == "" ? "_" : "" + cells[row.Index, 21].Value);
                            //醫學分數
                            data.medical_score = ("" + cells[row.Index, 22].Value == "" ? "_" : "" + cells[row.Index, 22].Value);
                            //生資食科分數
                            data.food_science_score = ("" + cells[row.Index, 23].Value == "" ? "_" : "" + cells[row.Index, 23].Value);
                            //地球環境分數
                            data.earth_enviroment_score = ("" + cells[row.Index, 24].Value == "" ? "_" : "" + cells[row.Index, 24].Value);
                            //藝術分數
                            data.art_score = ("" + cells[row.Index, 25].Value == "" ? "_" : "" + cells[row.Index, 25].Value);
                            //歷史文化分數
                            data.history_score = ("" + cells[row.Index, 26].Value == "" ? "_" : "" + cells[row.Index, 26].Value);
                            //傳播媒體分數
                            data.media_score = ("" + cells[row.Index, 27].Value == "" ? "_" : "" + cells[row.Index, 27].Value);
                            //教育訓練分數
                            data.education_scorescore = ("" + cells[row.Index, 28].Value == "" ? "_" : "" + cells[row.Index, 28].Value);
                            //心理學分數
                            data.psychology_score = ("" + cells[row.Index, 29].Value == "" ? "_" : "" + cells[row.Index, 29].Value);
                            //社會人類分數
                            data.society_score = ("" + cells[row.Index, 30].Value == "" ? "_" : "" + cells[row.Index, 30].Value);
                            //哲學宗教分數
                            data.philosophy_score = ("" + cells[row.Index, 31].Value == "" ? "_" : "" + cells[row.Index, 31].Value);
                            //治療諮商分數
                            data.consultation_score = ("" + cells[row.Index, 32].Value == "" ? "_" : "" + cells[row.Index, 32].Value);
                            //語文文學分數
                            data.language_score = ("" + cells[row.Index, 33].Value == "" ? "_" : "" + cells[row.Index, 33].Value);
                            //外國語文分數
                            data.foreign_language_score = ("" + cells[row.Index, 34].Value == "" ? "_" : "" + cells[row.Index, 34].Value);
                            //人力資源分數
                            data.human_resource_score = ("" + cells[row.Index, 35].Value == "" ? "_" : "" + cells[row.Index, 35].Value);
                            //顧客服務分數
                            data.customer_service_score = ("" + cells[row.Index, 36].Value == "" ? "_" : "" + cells[row.Index, 36].Value);
                            //管理分數
                            data.management_score = ("" + cells[row.Index, 37].Value == "" ? "_" : "" + cells[row.Index, 37].Value);
                            //銷售行銷分數
                            data.market_score = ("" + cells[row.Index, 38].Value == "" ? "_" : "" + cells[row.Index, 38].Value);
                            //經濟會計分數
                            data.accounting_score = ("" + cells[row.Index, 39].Value == "" ? "_" : "" + cells[row.Index, 39].Value);
                            //法律政治分數
                            data.law_score = ("" + cells[row.Index, 40].Value == "" ? "_" : "" + cells[row.Index, 40].Value);
                            //行政分數
                            data.administrative_score = ("" + cells[row.Index, 41].Value == "" ? "_" : "" + cells[row.Index, 41].Value);
                            

                            //// 學生ID
                            data.StudentID = studentID;

                            // 送出日期
                            data.ImplementationDate = dt;

                            #endregion

                            // 將 data 加入 list                                                                                        
                            interest_Test_Data_List.Add(data);
                        }
                        else
                        {
                            DAO.UDT_MajorIn_Test_Data_Def data = dataList[0];

                            //志願一
                            data.wish_1 = ("" + cells[row.Index, 4].Value == "" ? "_" : "" + cells[row.Index, 4].Value);
                            //志願二
                            data.wish_2 = ("" + cells[row.Index, 5].Value == "" ? "_" : "" + cells[row.Index, 5].Value);
                            //志願三
                            data.wish_3 = ("" + cells[row.Index, 6].Value == "" ? "_" : "" + cells[row.Index, 6].Value);
                            //志願四
                            data.wish_4 = ("" + cells[row.Index, 7].Value == "" ? "_" : "" + cells[row.Index, 7].Value);
                            //志願五
                            data.wish_5 = ("" + cells[row.Index, 8].Value == "" ? "_" : "" + cells[row.Index, 8].Value);
                            //志願六
                            data.wish_6 = ("" + cells[row.Index, 9].Value == "" ? "_" : "" + cells[row.Index, 9].Value);
                            //志願七
                            data.wish_7 = ("" + cells[row.Index, 10].Value == "" ? "_" : "" + cells[row.Index, 10].Value);
                            //志願八
                            data.wish_8 = ("" + cells[row.Index, 11].Value == "" ? "_" : "" + cells[row.Index, 11].Value);
                            //數學分數
                            data.math_score = ("" + cells[row.Index, 12].Value == "" ? "_" : "" + cells[row.Index, 12].Value);
                            //物理分數
                            data.phycics_score = ("" + cells[row.Index, 13].Value == "" ? "_" : "" + cells[row.Index, 13].Value);
                            //化學分數
                            data.chemistry_score = ("" + cells[row.Index, 14].Value == "" ? "_" : "" + cells[row.Index, 14].Value);
                            //資訊電子分數
                            data.information_electronics_score = ("" + cells[row.Index, 15].Value == "" ? "_" : "" + cells[row.Index, 15].Value);
                            //通訊電信分數
                            data.communication_telecommunications = ("" + cells[row.Index, 16].Value == "" ? "_" : "" + cells[row.Index, 16].Value);
                            //工程科技分數
                            data.engineer_technology_score = ("" + cells[row.Index, 17].Value == "" ? "_" : "" + cells[row.Index, 17].Value);
                            //機械分數
                            data.mechanism_score = ("" + cells[row.Index, 18].Value == "" ? "_" : "" + cells[row.Index, 18].Value);
                            //建築營造分數
                            data.building_construction_score = ("" + cells[row.Index, 19].Value == "" ? "_" : "" + cells[row.Index, 19].Value);
                            //設計分數
                            data.design_score = ("" + cells[row.Index, 20].Value == "" ? "_" : "" + cells[row.Index, 20].Value);
                            //生命科學分數
                            data.life_science_score = ("" + cells[row.Index, 21].Value == "" ? "_" : "" + cells[row.Index, 21].Value);
                            //醫學分數
                            data.medical_score = ("" + cells[row.Index, 22].Value == "" ? "_" : "" + cells[row.Index, 22].Value);
                            //生資食科分數
                            data.food_science_score = ("" + cells[row.Index, 23].Value == "" ? "_" : "" + cells[row.Index, 23].Value);
                            //地球環境分數
                            data.earth_enviroment_score = ("" + cells[row.Index, 24].Value == "" ? "_" : "" + cells[row.Index, 24].Value);
                            //藝術分數
                            data.art_score = ("" + cells[row.Index, 25].Value == "" ? "_" : "" + cells[row.Index, 25].Value);
                            //歷史文化分數
                            data.history_score = ("" + cells[row.Index, 26].Value == "" ? "_" : "" + cells[row.Index, 26].Value);
                            //傳播媒體分數
                            data.media_score = ("" + cells[row.Index, 27].Value == "" ? "_" : "" + cells[row.Index, 27].Value);
                            //教育訓練分數
                            data.education_scorescore = ("" + cells[row.Index, 28].Value == "" ? "_" : "" + cells[row.Index, 28].Value);
                            //心理學分數
                            data.psychology_score = ("" + cells[row.Index, 29].Value == "" ? "_" : "" + cells[row.Index, 29].Value);
                            //社會人類分數
                            data.society_score = ("" + cells[row.Index, 30].Value == "" ? "_" : "" + cells[row.Index, 30].Value);
                            //哲學宗教分數
                            data.philosophy_score = ("" + cells[row.Index, 31].Value == "" ? "_" : "" + cells[row.Index, 31].Value);
                            //治療諮商分數
                            data.consultation_score = ("" + cells[row.Index, 32].Value == "" ? "_" : "" + cells[row.Index, 32].Value);
                            //語文文學分數
                            data.language_score = ("" + cells[row.Index, 33].Value == "" ? "_" : "" + cells[row.Index, 33].Value);
                            //外國語文分數
                            data.foreign_language_score = ("" + cells[row.Index, 34].Value == "" ? "_" : "" + cells[row.Index, 34].Value);
                            //人力資源分數
                            data.human_resource_score = ("" + cells[row.Index, 35].Value == "" ? "_" : "" + cells[row.Index, 35].Value);
                            //顧客服務分數
                            data.customer_service_score = ("" + cells[row.Index, 36].Value == "" ? "_" : "" + cells[row.Index, 36].Value);
                            //管理分數
                            data.management_score = ("" + cells[row.Index, 37].Value == "" ? "_" : "" + cells[row.Index, 37].Value);
                            //銷售行銷分數
                            data.market_score = ("" + cells[row.Index, 38].Value == "" ? "_" : "" + cells[row.Index, 38].Value);
                            //經濟會計分數
                            data.accounting_score = ("" + cells[row.Index, 39].Value == "" ? "_" : "" + cells[row.Index, 39].Value);
                            //法律政治分數
                            data.law_score = ("" + cells[row.Index, 40].Value == "" ? "_" : "" + cells[row.Index, 40].Value);
                            //行政分數
                            data.administrative_score = ("" + cells[row.Index, 41].Value == "" ? "_" : "" + cells[row.Index, 41].Value);


                            //// 學生ID
                            data.StudentID = studentID;

                            // 送出日期
                            data.ImplementationDate = dt;

                            // 將 data 加入 list                                                                                        
                            interest_Test_Data_List.Add(data);
                        }
                    }
                    try
                    {
                        success_count++;

                        int progress = (success_count * 100 / ws.Cells.Rows.Count);

                        _bw.ReportProgress(progress);
                    }
                    catch (Exception ex)                    
                    {
                        MsgBox.Show(ex.Message, "錯誤", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);                                        
                    }
                    
                }

                //4.儲存上傳

                //  若仍有 錯誤資料，跳提醒，終止上傳。
                if (warningList.Count > 0)
                {
                    string warningMessages = "";

                    foreach (string warningmessage in warningList)
                    {
                        warningMessages += warningmessage;
                    }

                    MsgBox.Show(warningMessages, "錯誤", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);

                    return;
                }
                else
                {
                    // 一切無誤，儲存
                    interest_Test_Data_List.SaveAll();

                    // 儲存 log資料
                    DAO.LogTransfer _logTransfer = new DAO.LogTransfer();

                    StringBuilder logData = new StringBuilder();

                    logData.AppendLine("匯入年級:" + target_grade_year);
                    logData.AppendLine("匯入總人數:" + interest_Test_Data_List.Count);

                    _logTransfer.SaveLog("輔導系統.匯入興趣測驗資料", "匯入", "", "", logData);

                    MsgBox.Show("匯入成功", "完成", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
                }            
            }
            catch(Exception ex) 
            {
                // 針對Aspose 不支援 舊版 Excel 功能，另外在錯誤視窗提醒使用者重新存檔匯入
                if (ex.Source == "Aspose.Cells")
                {
                   MsgBox.Show(ex.Message+"EXCEL來源格式MicroSoft Excel5.0/95，過於老舊，本系統不支援匯入，請用Excel 將欲匯入檔案另存成MicroSoft Excel 97~2003 的.xls格式後，再使用該檔案匯入" , "錯誤", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                }
                else 
                {
                    MsgBox.Show(ex.Message, "錯誤", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);                
                }                
            }
        }

        void _worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("處理中...", e.ProgressPercentage);
        }

        private void _worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            FISCA.Presentation.MotherForm.SetStatusBarMessage("輔導系統-匯入興趣測驗資料完成");
            // 任務結束，關閉
            this.Close();            
        }

        // 關閉
        private void buttonX3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // 檢視 支援樣板
        private void labelX4_Click(object sender, EventArgs e)
        {
            Workbook wb = new Workbook(new MemoryStream(Properties.Resources.學系探索量表測驗匯入樣板));
            
            SaveFileDialog sd = new SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = "學系探索量表測驗匯入樣板.xlsx";
            sd.Filter = "Excel檔案 (*.xlsx)|*.xlsx|所有檔案 (*.*)|*.*";
            if (sd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    wb.Save(sd.FileName);
                    System.Diagnostics.Process.Start(sd.FileName);
                }
                catch
                {
                    MessageBox.Show("指定路徑無法存取。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);                    
                }
            }
        } 
    }
}



