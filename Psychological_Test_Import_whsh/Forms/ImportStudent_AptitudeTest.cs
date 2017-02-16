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
    public partial class ImportStudent_AptitudeTest: BaseForm
    {

        BackgroundWorker _bw = new BackgroundWorker();

        public string source_data = "";
        string target_grade_year = "";
        string date_time = "";
        DateTime dt;

        public ImportStudent_AptitudeTest()
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

                // 文件打開舊方法，已過時， 現在統一使用上面的new Workbook(textBoxX1.Text) 的方式，另外 請使用最新 Aspose.Cell_201402 ，可以避免很多讀取錯誤Bug
                //wb.Open(textBoxX1.Text, FileFormatType.Excel2007Xlsx);

                Worksheet ws = wb.Worksheets[0];

                Cells cells = ws.Cells;

                #region 驗證欄位 List
                List<string> verifiedDomainList1 = new List<string>();

                verifiedDomainList1.Add("語文推理");
                verifiedDomainList1.Add("數字推理");
                verifiedDomainList1.Add("圖形推理");
                verifiedDomainList1.Add("機械推理");
                verifiedDomainList1.Add("空間關係");
                verifiedDomainList1.Add("中文詞語");
                verifiedDomainList1.Add("英文詞語");
                verifiedDomainList1.Add("知覺速度");

                List<string> verifiedDomainList2 = new List<string>();

                verifiedDomainList2.Add("學業性向");
                verifiedDomainList2.Add("理工性向");
                verifiedDomainList2.Add("文科性向");

                List<string> verifiedDomainList3 = new List<string>();

                verifiedDomainList3.Add("知覺速度");

                List<string> verifiedDetailScorelList1 = new List<string>();

                verifiedDetailScorelList1.Add("原始分數");
                verifiedDetailScorelList1.Add("量表分數");
                verifiedDetailScorelList1.Add("百分等級");

                List<string> verifiedDetailScorelList2 = new List<string>();

                verifiedDetailScorelList2.Add("組合分數");
                verifiedDetailScorelList2.Add("百分等級");

                List<string> verifiedDetailScorelList3 = new List<string>();

                verifiedDetailScorelList3.Add("作答題數");
                verifiedDetailScorelList3.Add("答對題數");

                List<string> verifiedDetailScorelList4 = new List<string>();

                verifiedDetailScorelList4.Add("班級");
                verifiedDetailScorelList4.Add("座號");
                verifiedDetailScorelList4.Add("性別");
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

                #region 舊的寫死方法

                // 建立班級序號(在心測中心提供的匯入格式) 與 班級名稱的對照，班級序號 1 對應 班級名稱 一年01班 後續以此類推，
                // 文華一個年級 僅有20 班，扣掉 舞蹈班，實際測驗班級數 為 19 班 ，在此提供到30 個班級，供後續擴充
                // 原本是很想做一個及時偵測班級數目的對應Dict，但此專案是專做給文華的，先這樣子做比較簡單。

                //classNO_To_ClassName.Add("1", "一年01班");
                //classNO_To_ClassName.Add("2", "一年02班");
                //classNO_To_ClassName.Add("3", "一年03班");
                //classNO_To_ClassName.Add("4", "一年04班");
                //classNO_To_ClassName.Add("5", "一年05班");
                //classNO_To_ClassName.Add("6", "一年06班");
                //classNO_To_ClassName.Add("7", "一年07班");
                //classNO_To_ClassName.Add("8", "一年08班");
                //classNO_To_ClassName.Add("9", "一年09班");
                //classNO_To_ClassName.Add("10", "一年10班");
                //classNO_To_ClassName.Add("11", "一年11班");
                //classNO_To_ClassName.Add("12", "一年12班");
                //classNO_To_ClassName.Add("13", "一年13班");
                //classNO_To_ClassName.Add("14", "一年14班");
                //classNO_To_ClassName.Add("15", "一年15班");
                //classNO_To_ClassName.Add("16", "一年16班");
                //classNO_To_ClassName.Add("17", "一年17班");
                //classNO_To_ClassName.Add("18", "一年18班");
                //classNO_To_ClassName.Add("19", "一年19班");
                //classNO_To_ClassName.Add("20", "一年20班");
                //classNO_To_ClassName.Add("21", "一年21班");
                //classNO_To_ClassName.Add("22", "一年22班");
                //classNO_To_ClassName.Add("23", "一年23班");
                //classNO_To_ClassName.Add("24", "一年24班");
                //classNO_To_ClassName.Add("25", "一年25班");
                //classNO_To_ClassName.Add("26", "一年26班");
                //classNO_To_ClassName.Add("27", "一年27班");
                //classNO_To_ClassName.Add("28", "一年28班");
                //classNO_To_ClassName.Add("29", "一年29班");
                //classNO_To_ClassName.Add("30", "一年30班");  
                #endregion

                // 動態增加 對照
                int target_grade_class_list_index = 1;

                foreach (K12.Data.ClassRecord cr in target_grade_class_list)
                {
                    classNO_To_ClassName.Add("" + target_grade_class_list_index, cr.Name);

                    target_grade_class_list_index++;
                }
                #endregion

                // 建立班級、座號  對應 StudentID 對照表(使用 ClassIDs 來選取 本次範圍的學生)
                List<StudentRecord> allStudentList = K12.Data.Student.SelectByClassIDs(target_grade_classID_list);

                // 建立 班級、座號  對應 StudentID 之 dict  (因為本心理測驗施測對象 只有一年級)
                Dictionary<string, string> class_SeatNO_To_StudentID = new Dictionary<string, string>();

                foreach (StudentRecord sr in allStudentList)
                {
                    if (sr.Class != null)
                    {
                        // 同時具有班級名稱 與座號，才加入對照
                        if (sr.Class.Name != "" && "" + sr.SeatNo != "")
                        {
                            // key = 班級名稱_座號  ,EX: 一年01班_21號
                            string key = sr.Class.Name + "_" + sr.SeatNo;

                            class_SeatNO_To_StudentID.Add(key, sr.ID);
                        }
                    }
                }

                // 錯誤資料List
                List<string> errorList = new List<string>();

                // 1.驗證資料

                // 1.1 驗證欄位標題

                #region 驗證欄位標題
                int i = 0;

                foreach (string domain in verifiedDomainList1)
                {
                    if ("" + cells[3, (i * 3) + 3].Value != verifiedDomainList1[i])
                    {
                        errorList.Add("選擇檔案EXCEL 欄位於 第四列 第" + (i * 3 + 4) + "行處，領域與匯入格式不符，請確認");
                    }
                    if ("" + cells[3, (i * 3) + 4].Value != verifiedDomainList1[i])
                    {
                        errorList.Add("選擇檔案EXCEL 欄位於 第四列 第" + (i * 3 + 5) + "行處，領域與匯入格式不符，請確認");
                    }
                    if ("" + cells[3, (i * 3) + 5].Value != verifiedDomainList1[i])
                    {
                        errorList.Add("選擇檔案EXCEL 欄位於 第四列 第" + (i * 3 + 6) + "行處，領域與匯入格式不符，請確認");
                    }

                    i++;
                }


                i = 0;

                foreach (string domain in verifiedDomainList2)
                {
                    if ("" + cells[3, (i * 2) + 27].Value != verifiedDomainList2[i])
                    {
                        errorList.Add("選擇檔案EXCEL 欄位於 第四列 第" + (i * 2 + 28) + "行處，領域與匯入格式不符，請確認");
                    }
                    if ("" + cells[3, (i * 2) + 28].Value != verifiedDomainList2[i])
                    {
                        errorList.Add("選擇檔案EXCEL 欄位於 第四列 第" + (i * 2 + 29) + "行處，領域與匯入格式不符，請確認");
                    }
                    i++;
                }

                i = 0;

                foreach (string domain in verifiedDomainList3)
                {
                    if ("" + cells[3, (i * 2) + 33].Value != verifiedDomainList3[i])
                    {
                        errorList.Add("選擇檔案EXCEL 欄位於 第四列 第" + (i * 2 + 34) + "行處，領域與匯入格式不符，請確認");
                    }
                    if ("" + cells[3, (i * 2) + 34].Value != verifiedDomainList3[i])
                    {
                        errorList.Add("選擇檔案EXCEL 欄位於 第四列 第" + (i * 2 + 35) + "行處，領域與匯入格式不符，請確認");
                    }
                    i++;
                }

                for (int i_detail = 0; i_detail < 8; i_detail++)
                {
                    if ("" + cells[4, (i_detail * 3) + 3].Value != verifiedDetailScorelList1[0])
                    {
                        errorList.Add("選擇檔案EXCEL 欄位於 第五列 第" + (i_detail * 3 + 4) + "行處，領域與匯入格式不符，請確認");
                    }
                    if ("" + cells[4, (i_detail * 3) + 4].Value != verifiedDetailScorelList1[1])
                    {
                        errorList.Add("選擇檔案EXCEL 欄位於 第五列 第" + (i_detail * 3 + 5) + "行處，領域與匯入格式不符，請確認");
                    }
                    if ("" + cells[4, (i_detail * 3) + 5].Value != verifiedDetailScorelList1[2])
                    {
                        errorList.Add("選擇檔案EXCEL 欄位於 第五列 第" + (i_detail * 3 + 6) + "行處，領域與匯入格式不符，請確認");
                    }
                }

                for (int i_detail = 0; i_detail < 3; i_detail++)
                {
                    if ("" + cells[4, (i_detail * 2) + 27].Value != verifiedDetailScorelList2[0])
                    {
                        errorList.Add("選擇檔案EXCEL 欄位於 第五列 第" + (i_detail * 3 + 28) + "行處，領域與匯入格式不符，請確認");
                    }
                    if ("" + cells[4, (i_detail * 2) + 28].Value != verifiedDetailScorelList2[1])
                    {
                        errorList.Add("選擇檔案EXCEL 欄位於 第五列 第" + (i_detail * 3 + 29) + "行處，領域與匯入格式不符，請確認");
                    }
                }

                for (int i_detail = 0; i_detail < 1; i_detail++)
                {
                    if ("" + cells[4, (i_detail * 2) + 33].Value != verifiedDetailScorelList3[0])
                    {
                        errorList.Add("選擇檔案EXCEL 欄位於 第五列 第" + (i_detail * 3 + 34) + "行處，領域與匯入格式不符，請確認");
                    }
                    if ("" + cells[4, (i_detail * 2) + 34].Value != verifiedDetailScorelList3[1])
                    {
                        errorList.Add("選擇檔案EXCEL 欄位於 第五列 第" + (i_detail * 3 + 35) + "行處，領域與匯入格式不符，請確認");
                    }
                }

                if ("" + cells[4, 0].Value != verifiedDetailScorelList4[0])
                {
                    errorList.Add("選擇檔案EXCEL 欄位於 第五列 第" + 1 + "行處，KEY值與匯入格式不符，請確認");
                }
                if ("" + cells[4, 1].Value != verifiedDetailScorelList4[1])
                {
                    errorList.Add("選擇檔案EXCEL 欄位於 第五列 第" + 2 + "行處，KEY值與匯入格式不符，請確認");
                }
                if ("" + cells[4, 2].Value != verifiedDetailScorelList4[2])
                {
                    errorList.Add("選擇檔案EXCEL 欄位於 第五列 第" + 3 + "行處，KEY值與匯入格式不符，請確認");
                }

                #endregion


                //1.2  驗證班級學生資料

                #region 驗證班級學生資料

                Dictionary<string, List<string>> studentDict_By_Class = new Dictionary<string, List<string>>();

                foreach (var row in ws.Cells.Rows)
                {
                    if (row.Index >= 5 && "" + cells[row.Index, 0].Value != "" && "" + cells[row.Index, 1].Value != "")
                    {
                        string class_name = "";

                        string student_seat_number = "";

                        class_name = "" + cells[row.Index, 0].Value;

                        student_seat_number = "" + cells[row.Index, 1].Value;

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
                                errorList.Add("選擇檔案EXCEL 欄位  班級" + class_name + "座號" + student_seat_number + "資料重覆");
                            }
                        }
                    }
                    if (row.Index >= 5 && "" + cells[row.Index, 0].Value == "" && !row.IsBlank)
                    {
                        errorList.Add("選擇檔案EXCEL 欄位 於 第" + row.Index + "列 第" + 1 + "行處，沒有班級");
                    }
                    if (row.Index >= 5 && "" + cells[row.Index, 1].Value == "" && !row.IsBlank)
                    {
                        errorList.Add("選擇檔案EXCEL 欄位 於 第" + row.Index + "列 第" + 2 + "行處，沒有座號");
                    }
                    if (row.Index >= 5 && "" + cells[row.Index, 0].Value != "" && !row.IsBlank)
                    {
                        if (!classNO_To_ClassName.ContainsKey("" + cells[row.Index, 0].Value))
                        {
                            errorList.Add("選擇檔案EXCEL 欄位 於 第" + row.Index + "列 第" + 1 + "行處，不存在此班級編號，目前僅支援1~30班級編號。");
                        }

                    }

                    if (row.Index >= 5 && !row.IsBlank)
                    {
                        string key = "";

                        // 取得學生ID Key, key = 班級名稱_座號  ,EX: 一年01班_21號
                        if (classNO_To_ClassName.ContainsKey("" + cells[row.Index, 0].Value))
                        {
                            key = classNO_To_ClassName["" + cells[row.Index, 0].Value] + "_" + cells[row.Index, 1].Value;
                        }

                        if (!class_SeatNO_To_StudentID.ContainsKey(key))
                        {
                            errorList.Add("第" + (row.Index + 1) + "行，該班級座號的學生並不存在於本系統中，請檢察");
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
                List<DAO.UDT_Aptitude_Test_Data_Def> aptitude_Test_Data_List = new List<DAO.UDT_Aptitude_Test_Data_Def>();

                // 警告資料List
                List<string> warningList = new List<string>();

                FISCA.UDT.AccessHelper accesshelper = new AccessHelper();

                int success_count = 0;

                foreach (var row in ws.Cells.Rows)
                {
                    string key = "";
                    string studentID = "";

                    if (row.Index >= 5 && !row.IsBlank)
                    {
                        // 取得學生ID Key, key = 班級名稱_座號  ,EX: 一年01班_21號
                        if (classNO_To_ClassName.ContainsKey("" + cells[row.Index, 0].Value))
                        {
                            key = classNO_To_ClassName["" + cells[row.Index, 0].Value] + "_" + cells[row.Index, 1].Value;
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


                        List<DAO.UDT_Aptitude_Test_Data_Def> dataList = accesshelper.Select<DAO.UDT_Aptitude_Test_Data_Def>("ref_student_id =" + "'" + studentID + "'");

                        if (dataList.Count == 0)
                        {
                            DAO.UDT_Aptitude_Test_Data_Def data = new DAO.UDT_Aptitude_Test_Data_Def();

                            #region 填值
                            //語文推理 : 原始分數、量表分數、百分等級
                            data.language_reasoning_original_score = "" + cells[row.Index, 3].Value;
                            data.language_reasoning_scale_score = "" + cells[row.Index, 4].Value;
                            data.language_reasoning_pr_level = "" + cells[row.Index, 5].Value;

                            //數字推理 : 原始分數、量表分數、百分等級
                            data.digit_reasoning_original_score = "" + cells[row.Index, 6].Value;
                            data.digit_reasoning_scale_score = "" + cells[row.Index, 7].Value;
                            data.digit_reasoning_pr_level = "" + cells[row.Index, 8].Value;

                            //圖形推理 : 原始分數、量表分數、百分等級
                            data.image_reasoning_original_score = "" + cells[row.Index, 9].Value;
                            data.image_reasoning_scale_score = "" + cells[row.Index, 10].Value;
                            data.image_reasoning_pr_level = "" + cells[row.Index, 11].Value;

                            //機械推理 : 原始分數、量表分數、百分等級
                            data.mechanical_reasoning_original_score = "" + cells[row.Index, 12].Value;
                            data.mechanical_reasoning_scale_score = "" + cells[row.Index, 13].Value;
                            data.mechanical_reasoning_pr_level = "" + cells[row.Index, 14].Value;

                            //空間關係 : 原始分數、量表分數、百分等級
                            data.dimension_relation_original_score = "" + cells[row.Index, 15].Value;
                            data.dimension_relation_scale_score = "" + cells[row.Index, 16].Value;
                            data.dimension_relation_pr_level = "" + cells[row.Index, 17].Value;

                            //中文詞語 : 原始分數、量表分數、百分等級
                            data.chineses_words_original_score = "" + cells[row.Index, 18].Value;
                            data.chineses_words_scale_score = "" + cells[row.Index, 19].Value;
                            data.chineses_words_pr_level = "" + cells[row.Index, 20].Value;

                            //英文詞語 : 原始分數、量表分數、百分等級
                            data.english_words_original_score = "" + cells[row.Index, 21].Value;
                            data.english_words_scale_score = "" + cells[row.Index, 22].Value;
                            data.english_words_pr_level = "" + cells[row.Index, 23].Value;

                            //知覺速度 : 原始分數、量表分數、百分等級
                            data.perception_time_original_score = "" + cells[row.Index, 24].Value;
                            data.perception_time_scale_score = "" + cells[row.Index, 25].Value;
                            data.perception_time_pr_level = "" + cells[row.Index, 26].Value;

                            //學業性向 : 組合分數、百分等級
                            data.learning_aptitude_assemble_score = "" + cells[row.Index, 27].Value;
                            data.learning_aptitude_pr_level = "" + cells[row.Index, 28].Value;

                            //理工性向 : 組合分數、百分等級
                            data.science_aptitude_assemble_score = "" + cells[row.Index, 29].Value;
                            data.science_aptitude_pr_level = "" + cells[row.Index, 30].Value;

                            //文科性向 : 組合分數、百分等級
                            data.literal_aptitude_assemble_score = "" + cells[row.Index, 31].Value;
                            data.literal_aptitude_pr_level = "" + cells[row.Index, 32].Value;

                            //知覺速度 : 作答題數、答對題數
                            data.perception_time_complete_quiz_count = "" + cells[row.Index, 33].Value;
                            data.perception_time_correct_quiz_count = "" + cells[row.Index, 34].Value;

                            //// 學生ID
                            data.StudentID = studentID;

                            // 送出日期
                            data.ImplementationDate = dt;
                            #endregion

                            // 將 data 加入 list                                                                                        
                            aptitude_Test_Data_List.Add(data);

                        }
                        else
                        {
                            DAO.UDT_Aptitude_Test_Data_Def data = dataList[0];

                            #region 填值
                            //語文推理 : 原始分數、量表分數、百分等級
                            data.language_reasoning_original_score = "" + cells[row.Index, 3].Value;
                            data.language_reasoning_scale_score = "" + cells[row.Index, 4].Value;
                            data.language_reasoning_pr_level = "" + cells[row.Index, 5].Value;

                            //數字推理 : 原始分數、量表分數、百分等級
                            data.digit_reasoning_original_score = "" + cells[row.Index, 6].Value;
                            data.digit_reasoning_scale_score = "" + cells[row.Index, 7].Value;
                            data.digit_reasoning_pr_level = "" + cells[row.Index, 8].Value;

                            //圖形推理 : 原始分數、量表分數、百分等級
                            data.image_reasoning_original_score = "" + cells[row.Index, 9].Value;
                            data.image_reasoning_scale_score = "" + cells[row.Index, 10].Value;
                            data.image_reasoning_pr_level = "" + cells[row.Index, 11].Value;

                            //機械推理 : 原始分數、量表分數、百分等級
                            data.mechanical_reasoning_original_score = "" + cells[row.Index, 12].Value;
                            data.mechanical_reasoning_scale_score = "" + cells[row.Index, 13].Value;
                            data.mechanical_reasoning_pr_level = "" + cells[row.Index, 14].Value;

                            //空間關係 : 原始分數、量表分數、百分等級
                            data.dimension_relation_original_score = "" + cells[row.Index, 15].Value;
                            data.dimension_relation_scale_score = "" + cells[row.Index, 16].Value;
                            data.dimension_relation_pr_level = "" + cells[row.Index, 17].Value;

                            //中文詞語 : 原始分數、量表分數、百分等級
                            data.chineses_words_original_score = "" + cells[row.Index, 18].Value;
                            data.chineses_words_scale_score = "" + cells[row.Index, 19].Value;
                            data.chineses_words_pr_level = "" + cells[row.Index, 20].Value;

                            //英文詞語 : 原始分數、量表分數、百分等級
                            data.english_words_original_score = "" + cells[row.Index, 21].Value;
                            data.english_words_scale_score = "" + cells[row.Index, 22].Value;
                            data.english_words_pr_level = "" + cells[row.Index, 23].Value;

                            //知覺速度 : 原始分數、量表分數、百分等級
                            data.perception_time_original_score = "" + cells[row.Index, 24].Value;
                            data.perception_time_scale_score = "" + cells[row.Index, 25].Value;
                            data.perception_time_pr_level = "" + cells[row.Index, 26].Value;

                            //學業性向 : 組合分數、百分等級
                            data.learning_aptitude_assemble_score = "" + cells[row.Index, 27].Value;
                            data.learning_aptitude_pr_level = "" + cells[row.Index, 28].Value;

                            //理工性向 : 組合分數、百分等級
                            data.science_aptitude_assemble_score = "" + cells[row.Index, 29].Value;
                            data.science_aptitude_pr_level = "" + cells[row.Index, 30].Value;

                            //文科性向 : 組合分數、百分等級
                            data.literal_aptitude_assemble_score = "" + cells[row.Index, 31].Value;
                            data.literal_aptitude_pr_level = "" + cells[row.Index, 32].Value;

                            //知覺速度 : 作答題數、答對題數
                            data.perception_time_complete_quiz_count = "" + cells[row.Index, 33].Value;
                            data.perception_time_correct_quiz_count = "" + cells[row.Index, 34].Value;

                            //// 學生ID
                            data.StudentID = studentID;

                            // 送出日期
                            data.ImplementationDate = dt;
                            #endregion

                            // 將 data 加入 list                                                                                        
                            aptitude_Test_Data_List.Add(data);
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
                    aptitude_Test_Data_List.SaveAll();

                    // 儲存 log資料
                    DAO.LogTransfer _logTransfer = new DAO.LogTransfer();

                    StringBuilder logData = new StringBuilder();

                    logData.AppendLine("匯入年級:" + target_grade_year);
                    logData.AppendLine("匯入總人數:" + aptitude_Test_Data_List.Count);

                    _logTransfer.SaveLog("輔導系統.匯入性向測驗資料", "匯入", "", "", logData);

                    MsgBox.Show("匯入成功", "完成", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
                }
            
            }
            catch(Exception ex) 
            {
                MsgBox.Show(ex.Message, "錯誤", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
            }
        }

        void _worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("處理中...", e.ProgressPercentage);
        }


        private void _worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            FISCA.Presentation.MotherForm.SetStatusBarMessage("輔導系統-匯入性向測驗資料完成");
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
            Workbook wb = new Workbook(new MemoryStream(Properties.Resources.新編多元性向測驗文華高中測驗範例樣板_CSV));
            
            SaveFileDialog sd = new SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = "新編多元性向測驗文華高中測驗範例樣板.xlsx";
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

