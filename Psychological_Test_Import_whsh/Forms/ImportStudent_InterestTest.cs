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
    public partial class ImportStudent_InterestTest: BaseForm
    {

        BackgroundWorker _bw = new BackgroundWorker();

        public string source_data = "";
        string target_grade_year = "";
        string date_time = "";
        DateTime dt;
        bool useIDNumberCheck = false;
        

        public ImportStudent_InterestTest()
        {
            InitializeComponent();

            comboBoxEx1.Items.Add("1");
            comboBoxEx1.Items.Add("2");
            comboBoxEx1.Items.Add("3");

            // 預設為一年級
            comboBoxEx1.SelectedIndex = 0;            
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
            target_grade_year = comboBoxEx1.Text;
            date_time = dateTimeInput1.Text;
            dt = dateTimeInput1.Value;

            useIDNumberCheck = checkBox1.Checked;

            // 若沒選取來源檔案，中止程序
            if (source_data == "")
            {
                MsgBox.Show("請選擇來源檔案", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            // 若沒選取匯入年級，中止程序
            if (target_grade_year == "")
            {
                MsgBox.Show("請選擇年級", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            // 若沒選取送出日期，中止程序
            if (""+dateTimeInput1.Text == "")
            {
                MsgBox.Show("請選擇送出日期", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 暫停UI 功能
            textBoxX1.Enabled = false;
            comboBoxEx1.Enabled = false;
            dateTimeInput1.Enabled = false;
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

                verifiedDomainList1.Add("單位識別碼");
                verifiedDomainList1.Add("班級代碼");
                verifiedDomainList1.Add("班級名稱");
                verifiedDomainList1.Add("座號");
                verifiedDomainList1.Add("身分證號");
                verifiedDomainList1.Add("抓週第一碼");
                verifiedDomainList1.Add("抓週第二碼");
                verifiedDomainList1.Add("抓週第三碼");
                verifiedDomainList1.Add("興趣第一碼");
                verifiedDomainList1.Add("興趣第二碼");
                verifiedDomainList1.Add("興趣第三碼");
                verifiedDomainList1.Add("R型總分");
                verifiedDomainList1.Add("I型總分");
                verifiedDomainList1.Add("A型總分");
                verifiedDomainList1.Add("S型總分");
                verifiedDomainList1.Add("E型總分");
                verifiedDomainList1.Add("C型總分");
                verifiedDomainList1.Add("興趣代碼");
                //verifiedDomainList1.Add("扇型區域");
                verifiedDomainList1.Add("諧和度");
                verifiedDomainList1.Add("區分值");
                //verifiedDomainList1.Add("一致性");
      
                #endregion

                // 全部班級list
                List<K12.Data.ClassRecord> all_class_list = K12.Data.Class.SelectAll();

                // 目標年級班級  list
                List<K12.Data.ClassRecord> target_grade_class_list = new List<ClassRecord>();

                // 目標年級班級ID  list (之後作為 Select studentRecord 使用)
                List<string> target_grade_classID_list = new List<string>();

                foreach (K12.Data.ClassRecord cr in all_class_list)
                {
                    if ("" + cr.GradeYear == target_grade_year)
                    {
                        target_grade_class_list.Add(cr);

                        target_grade_classID_list.Add(cr.ID);
                    }
                }

                // 將target_grade_class_list 排序，EX:一年01班、一年02班、一年03班......
                target_grade_class_list.Sort();

                #region 班級序號、班級名稱 對照
                Dictionary<string, string> classNO_To_ClassName = new Dictionary<string, string>();
                
                // 動態增加 對照
                int target_grade_class_list_index = 1;

                foreach (K12.Data.ClassRecord cr in target_grade_class_list)
                {
                    string key = "" + cr.GradeYear + (target_grade_class_list_index < 10 ? "0" : "") + target_grade_class_list_index;

                    classNO_To_ClassName.Add(key, cr.Name);

                    target_grade_class_list_index++;
                }
                #endregion

                // 錯誤資料List
                List<string> errorList = new List<string>();

                // 建立班級、座號  對應 StudentID 對照表(使用 ClassIDs 來選取 本次範圍的學生)
                List<StudentRecord> allStudentList = K12.Data.Student.SelectByClassIDs(target_grade_classID_list);

                // 建立 班級、座號  對應 StudentID 之 dict  (因為本心理測驗施測對象 只有一年級)
                Dictionary<string, string> class_SeatNO_To_StudentID = new Dictionary<string, string>();

                // 建立 班級、座號  對應 身分證字號 IDNumber 之 dict
                Dictionary<string, string> class_SeatNO_To_IDNumber = new Dictionary<string, string>();

                foreach (StudentRecord sr in allStudentList)
                {
                    if (sr.Class != null && (sr.Status == StudentRecord.StudentStatus.一般 |sr.Status == StudentRecord.StudentStatus.延修))
                    {
                        // 同時具有班級名稱 與座號，才加入對照
                        if (sr.Class.Name != "" && "" + sr.SeatNo != "")
                        {
                            // key = 班級名稱_座號  ,EX: 一年01班_21號
                            string key = ""+sr.Class.Name + "_" + sr.SeatNo;

                            if (!class_SeatNO_To_StudentID.ContainsKey(key))
                            {
                                class_SeatNO_To_StudentID.Add(key, sr.ID);
                            }
                            else 
                            {
                                errorList.Add("班級:" + sr.Class.Name + "座號:" + sr.SeatNo+"在系統中具有兩筆資料，將導致無法分辨確切學生身分，請確認檢查該資料。");                            
                            }

                            if (useIDNumberCheck && !class_SeatNO_To_IDNumber.ContainsKey(key))
                            {
                                if (sr.IDNumber != "" && sr.IDNumber != null)
                                {
                                    class_SeatNO_To_IDNumber.Add(key, sr.IDNumber);
                                }
                                else 
                                {
                                    errorList.Add("班級:" + sr.Class.Name + "座號:" + sr.SeatNo + "在系統中其學生資料上沒有身分證字號資料，將導致無法驗證學生身分，請在系統補齊其資料。");                                                            
                                }                                
                            }
                            else 
                            {
                                // 如果有重覆的學生資料，上面class_SeatNO_To_StudentID 的整理就做一遍了，不需要存兩筆一樣的資訊。                            
                            }                            
                        }
                    }
                }
                
                // 1.驗證資料

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
                    if (row.Index>0 && "" + cells[row.Index, 2].Value != "" && "" + cells[row.Index, 3].Value != "")
                    {
                        string class_name = "";

                        string student_seat_number = "";

                        class_name = "" + cells[row.Index, 2].Value;

                        student_seat_number = "" + cells[row.Index, 3].Value;

                        if (!studentDict_By_Class.ContainsKey(class_name))
                        {
                            studentDict_By_Class.Add(class_name, new List<string>());

                            studentDict_By_Class[class_name].Add(student_seat_number);
                        }
                        else
                        {
                            if (!studentDict_By_Class[class_name].Contains(student_seat_number))
                            {
                                studentDict_By_Class[class_name].Add(student_seat_number);
                            }
                            else
                            {
                                errorList.Add("第" + (row.Index + 1) + "列"+ "班級:"  + class_name + "座號" + student_seat_number + "資料重覆");
                            }
                        }
                    }
                    if (row.Index > 0 && "" + cells[row.Index, 2].Value == "" && !row.IsBlank)
                    {
                        errorList.Add("第" + (row.Index + 1) + "列 第" + 2 + "行處，沒有班級");
                    }
                    if (row.Index > 0 && "" + cells[row.Index, 3].Value == "" && !row.IsBlank)
                    {
                        errorList.Add("第" + (row.Index + 1) + "列 第" + 3 + "行處，沒有座號");
                    }
                    if (row.Index > 0 && "" + cells[row.Index, 2].Value != "" && !row.IsBlank)
                    {
                        if (!classNO_To_ClassName.ContainsKey("" + cells[row.Index, 2].Value))
                        {
                            errorList.Add("第" + (row.Index + 1) + "列 ，不存在此班級編號");
                        }
                    }
                    // 身分證字號 與班級座號座的驗證
                    if (useIDNumberCheck && row.Index > 0 && "" + cells[row.Index, 2].Value != "" && "" + cells[row.Index, 3].Value != "" && "" + cells[row.Index, 4].Value != "" && !row.IsBlank)
                    {
                        if (classNO_To_ClassName.ContainsKey("" + cells[row.Index, 2].Value))
                        {
                            string key = "" + classNO_To_ClassName["" + cells[row.Index, 2].Value] + "_" + cells[row.Index, 3].Value;

                            if (!class_SeatNO_To_IDNumber.ContainsKey(key))
                            {                                
                                errorList.Add("第" + (row.Index + 1) + "列 第" + 5 + "行處，在身分證字號驗證無法找到對照，請檢察其班級、座號、身分證字號輸入格式是否與他人不相同。");
                            }
                            else
                            {
                                if ("" + cells[row.Index, 4].Value != class_SeatNO_To_IDNumber[key])
                                {
                                    errorList.Add("第" + (row.Index + 1) + "列 第" + 5 + "行處，該學生身分證字號與系統中不同，請檢察。");
                                }
                            }
                        }
                        else 
                        {
                            // 如果 班及編號不存在 上面classNO_To_ClassName 的驗證就會加入 errorlist，不再重覆做                                                
                        }
                    }

                    if (row.Index > 0 && !row.IsBlank)
                    {
                        string key = "";

                        // 取得學生ID Key, key = 班級名稱_座號  ,EX: 一年01班_21號
                        if (classNO_To_ClassName.ContainsKey("" + cells[row.Index, 2].Value))
                        {
                            key = classNO_To_ClassName["" + cells[row.Index, 2].Value] + "_" + cells[row.Index, 3].Value;
                        }
                        if (!class_SeatNO_To_StudentID.ContainsKey(key))
                        {
                            errorList.Add("第" + (row.Index + 1) + "列，該班級座號的學生並不存在於本系統中，請檢察");
                        }
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
                List<DAO.UDT_Interest_Test_Data_Def_2018_09> interest_Test_Data_List = new List<DAO.UDT_Interest_Test_Data_Def_2018_09>();

                // 警告資料List
                List<string> warningList = new List<string>();

                FISCA.UDT.AccessHelper accesshelper = new AccessHelper();

                int success_count = 0;

                foreach (var row in ws.Cells.Rows)
                {
                    string key = "";
                    string studentID = "";

                    if (row.Index > 0 && !row.IsBlank)
                    {
                        // 取得學生ID Key, key = 班級名稱_座號  ,EX: 一年01班_01號
                        if (classNO_To_ClassName.ContainsKey("" + cells[row.Index, 2].Value))
                        {
                            key = classNO_To_ClassName["" + cells[row.Index, 2].Value] + "_" + cells[row.Index, 3].Value;
                        }

                        if (class_SeatNO_To_StudentID.ContainsKey(key))
                        {
                            // 學生ID
                            studentID = "" + int.Parse(class_SeatNO_To_StudentID[key]);
                        }
                        else
                        {
                            warningList.Add("第" + (row.Index + 1) + "行，該班級座號的學生並不存在於本系統中，請檢察");
                        }
                        List<DAO.UDT_Interest_Test_Data_Def_2018_09> dataList = accesshelper.Select<DAO.UDT_Interest_Test_Data_Def_2018_09>("ref_student_id =" + "'" + studentID + "'");

                        if (dataList.Count == 0)
                        {
                            DAO.UDT_Interest_Test_Data_Def_2018_09 data = new DAO.UDT_Interest_Test_Data_Def_2018_09();

                            #region 填值

                            //抓週第一碼
                            data.gift_test_first_code = ("" + cells[row.Index, 5].Value == "" ? "_" : "" + cells[row.Index, 5].Value);
                            //抓週第二碼
                            data.gift_test_second_code = ("" + cells[row.Index, 6].Value == "" ? "_" : "" + cells[row.Index, 6].Value);
                            //抓週第三碼
                            data.gift_test_third_code = ("" + cells[row.Index, 7].Value == "" ? "_" : "" + cells[row.Index, 7].Value);
                            //興趣第一碼
                            data.interest_first_code = ("" + cells[row.Index, 8].Value == "" ? "_" : "" + cells[row.Index, 8].Value);
                            //興趣第二碼
                            data.interest_second_code = ("" + cells[row.Index, 9].Value == "" ? "_" : "" + cells[row.Index, 9].Value);
                            //興趣第三碼
                            data.interest_third_code = ("" + cells[row.Index, 10].Value == "" ? "_" : "" + cells[row.Index, 10].Value);
                            //R型總分
                            data.r_type_score = ("" + cells[row.Index, 11].Value == "" ? "_" : "" + cells[row.Index, 11].Value);
                            //I型總分
                            data.i_type_score = ("" + cells[row.Index, 12].Value == "" ? "_" : "" + cells[row.Index, 12].Value);
                            //A型總分
                            data.a_type_score = ("" + cells[row.Index, 13].Value == "" ? "_" : "" + cells[row.Index, 13].Value);
                            //S型總分
                            data.s_type_score = ("" + cells[row.Index, 14].Value == "" ? "_" : "" + cells[row.Index, 14].Value);
                            //E型總分
                            data.e_type_score = ("" + cells[row.Index, 15].Value == "" ? "_" : "" + cells[row.Index, 15].Value);
                            //C型總分
                            data.c_type_score = ("" + cells[row.Index, 16].Value == "" ? "_" : "" + cells[row.Index, 16].Value);
                            //興趣代碼
                            data.interest_code = ("" + cells[row.Index, 17].Value == "" ? "_" : "" + cells[row.Index, 17].Value);                            
                            //諧和度
                            data.coordinate_index = ("" + cells[row.Index, 18].Value == "" ? "_" : "" + cells[row.Index, 18].Value);
                            //區分值
                            data.distinguishing_index = ("" + cells[row.Index, 19].Value == "" ? "_" : "" + cells[row.Index, 19].Value);
                           
                                                                                                                
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
                            DAO.UDT_Interest_Test_Data_Def_2018_09 data = dataList[0];

                            #region 填值
                            //抓週第一碼
                            data.gift_test_first_code = ("" + cells[row.Index, 5].Value == "" ? "_" : "" + cells[row.Index, 5].Value);
                            //抓週第二碼
                            data.gift_test_second_code = ("" + cells[row.Index, 6].Value == "" ? "_" : "" + cells[row.Index, 6].Value);
                            //抓週第三碼
                            data.gift_test_third_code = ("" + cells[row.Index, 7].Value == "" ? "_" : "" + cells[row.Index, 7].Value);
                            //興趣第一碼
                            data.interest_first_code = ("" + cells[row.Index, 8].Value == "" ? "_" : "" + cells[row.Index, 8].Value);
                            //興趣第二碼
                            data.interest_second_code = ("" + cells[row.Index, 9].Value == "" ? "_" : "" + cells[row.Index, 9].Value);
                            //興趣第三碼
                            data.interest_third_code = ("" + cells[row.Index, 10].Value == "" ? "_" : "" + cells[row.Index, 10].Value);
                            //R型總分
                            data.r_type_score = ("" + cells[row.Index, 11].Value == "" ? "_" : "" + cells[row.Index, 11].Value);
                            //I型總分
                            data.i_type_score = ("" + cells[row.Index, 12].Value == "" ? "_" : "" + cells[row.Index, 12].Value);
                            //A型總分
                            data.a_type_score = ("" + cells[row.Index, 13].Value == "" ? "_" : "" + cells[row.Index, 13].Value);
                            //S型總分
                            data.s_type_score = ("" + cells[row.Index, 14].Value == "" ? "_" : "" + cells[row.Index, 14].Value);
                            //E型總分
                            data.e_type_score = ("" + cells[row.Index, 15].Value == "" ? "_" : "" + cells[row.Index, 15].Value);
                            //C型總分
                            data.c_type_score = ("" + cells[row.Index, 16].Value == "" ? "_" : "" + cells[row.Index, 16].Value);
                            //興趣代碼
                            data.interest_code = ("" + cells[row.Index, 17].Value == "" ? "_" : "" + cells[row.Index, 17].Value);                            
                            //諧和度
                            data.coordinate_index = ("" + cells[row.Index, 18].Value == "" ? "_" : "" + cells[row.Index, 18].Value);
                            //區分值
                            data.distinguishing_index = ("" + cells[row.Index, 19].Value == "" ? "_" : "" + cells[row.Index, 19].Value);
                            
                            
                            //// 學生ID
                            data.StudentID = studentID;

                            // 送出日期
                            data.ImplementationDate = dt;
                            #endregion

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
            Workbook wb = new Workbook(new MemoryStream(Properties.Resources._2017興趣量表新樣板格式));
            
            SaveFileDialog sd = new SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = "興趣測驗範例樣板.xlsx";
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


