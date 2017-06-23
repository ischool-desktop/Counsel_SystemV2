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


namespace AfterGratuationStatics_whsh.Forms
{
    public partial class Step3_Form : BaseForm
    {
        BackgroundWorker _bw = new BackgroundWorker();

        // 檔案位置
        public string source_data = "";

        public Step3_Form()
        {
            InitializeComponent();
        }

        // 取消
        private void buttonX3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //選擇檔案來源
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

        //匯入
        private void buttonX2_Click(object sender, EventArgs e)
        {
            source_data = textBoxX1.Text;

            // 錯誤資料List
            List<string> errorList = new List<string>();

            // 若沒選取來源檔案，中止程序
            if (source_data == "")
            {
                MsgBox.Show("請選擇來源檔案", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //載入匯入檔案
            Workbook wb = new Workbook(source_data);

            // 文件打開舊方法，已過時， 現在統一使用上面的new Workbook(textBoxX1.Text) 的方式，另外 請使用最新 Aspose.Cell_201402 ，可以避免很多讀取錯誤Bug
            //wb.Open(textBoxX1.Text, FileFormatType.Excel2007Xlsx);

            #region 資料驗證

            #region 驗證畢業生榜單

            // 檢查樣版格式是否正確
            if (wb.Worksheets["畢業生榜單"] == null)
            {
                MsgBox.Show("Excel檔案 沒有'畢業生榜單' 頁籤，請確認樣板格式正確", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Worksheet ws_graduated_list = wb.Worksheets["畢業生榜單"];

            Cells cells_graduated_list = ws_graduated_list.Cells;

            if ("" + cells_graduated_list[0, 0].Value != "學號")
            {
                errorList.Add("頁籤:畢業生榜單，第一行表頭應為:學號");
            }
            if ("" + cells_graduated_list[0, 1].Value != "科別")
            {
                errorList.Add("頁籤:畢業生榜單，第二行表頭應為:科別");
            }
            if ("" + cells_graduated_list[0, 2].Value != "姓名")
            {
                errorList.Add("頁籤:畢業生榜單，第三行表頭應為:姓名");
            }
            if ("" + cells_graduated_list[0, 3].Value != "性別")
            {
                errorList.Add("頁籤:畢業生榜單，第四行表頭應為:性別");
            }
            if ("" + cells_graduated_list[0, 4].Value != "班級")
            {
                errorList.Add("頁籤:畢業生榜單，第五行表頭應為:班級");
            }
            if ("" + cells_graduated_list[0, 5].Value != "座號")
            {
                errorList.Add("頁籤:畢業生榜單，第六行表頭應為:座號");
            }
            if ("" + cells_graduated_list[0, 6].Value != "學校")
            {
                errorList.Add("頁籤:畢業生榜單，第七行表頭應為:學校");
            }
            if ("" + cells_graduated_list[0, 7].Value != "學系")
            {
                errorList.Add("頁籤:畢業生榜單，第八行表頭應為:學系");
            }
            if ("" + cells_graduated_list[0, 8].Value != "學生系統分類")
            {
                errorList.Add("頁籤:畢業生榜單，第八行表頭應為:學生系統分類");
            }
            if ("" + cells_graduated_list[0, 9].Value != "人工指定分類")
            {
                errorList.Add("頁籤:畢業生榜單，第八行表頭應為:人工指定分類");
            }
            #endregion

            #region 驗證人工指定分類
            List<string> totalCategory_list = new List<string>();

            #region 分類清單
            totalCategory_list.Add("A01公立大學校院日間部(含第二部、四技)");
            totalCategory_list.Add("A02公立大學校院進修學士班(含四技)");
            totalCategory_list.Add("A03私立大學校院日間部(含第二部、四技)");
            totalCategory_list.Add("A04私立大學校院進修學士班(含四技)");
            totalCategory_list.Add("A05公立二專日間部");
            totalCategory_list.Add("A06公立二專進修(夜間)部(含附設進修專校)");
            totalCategory_list.Add("A07私立二專日間部");
            totalCategory_list.Add("A08私立二專進修(夜間)部(含附設進修專校)");
            totalCategory_list.Add("A09警察大學");
            totalCategory_list.Add("A10警察專科學校");
            totalCategory_list.Add("A11軍事院校");
            totalCategory_list.Add("A12赴國外、大陸就讀");
            totalCategory_list.Add("A99其他學校");
            totalCategory_list.Add("B01農、林、漁、牧業");
            totalCategory_list.Add("B02礦業及土石採取業");
            totalCategory_list.Add("B03製造業");
            totalCategory_list.Add("B04電力及燃氣供應業");
            totalCategory_list.Add("B05用水供應及污染整治業");
            totalCategory_list.Add("B06營建工程業");
            totalCategory_list.Add("B07批發及零售業");
            totalCategory_list.Add("B08運輸及倉儲業");
            totalCategory_list.Add("B09住宿及餐飲業");
            totalCategory_list.Add("B10出版、影音製作、傳播及資通訊服務業");
            totalCategory_list.Add("B11金融及保險業");
            totalCategory_list.Add("B12不動產業");
            totalCategory_list.Add("B13專業、科學及技術服務業");
            totalCategory_list.Add("B14支援服務業");
            totalCategory_list.Add("B15公共行政及國防；強制性社會安全");
            totalCategory_list.Add("B16教育業");
            totalCategory_list.Add("B17醫療保健及社會工作服務業");
            totalCategory_list.Add("B18藝術、娛樂及休閒服務業");
            totalCategory_list.Add("B99其他服務業");
            totalCategory_list.Add("C01正在接受職業訓練");
            totalCategory_list.Add("C02正在軍中服役");
            totalCategory_list.Add("C03需要工作而未找到");
            totalCategory_list.Add("C04正在補習或自修準備升學");
            totalCategory_list.Add("C05因健康不良在家休養");
            totalCategory_list.Add("C06準備出國");
            totalCategory_list.Add("C99其他");
            totalCategory_list.Add("D01遷居國外");
            totalCategory_list.Add("D02死亡");
            totalCategory_list.Add("D03無法聯繫或不詳");
            #endregion

            List<string> stuNumber_Name_list = new List<string>();

            // 驗證使用者輸入 人工指定分類 是否正確
            foreach (var row in cells_graduated_list.Rows)
            {
                if (row.Index > 0 && !row.IsBlank)
                {
                    if (("" + cells_graduated_list[row.Index, 9].Value) != "" && !totalCategory_list.Contains("" + cells_graduated_list[row.Index, 9].Value))
                    {
                        errorList.Add("頁籤:畢業生榜單，第" + (row.Index + 1) + "列，學生:" + cells_graduated_list[row.Index, 2].Value + " ，人工指定分類不存在，請確認輸入格式內容正確。");
                    }

                    //Key: 學號+ _ + 姓名
                    stuNumber_Name_list.Add("" + cells_graduated_list[row.Index, 0].Value + "_" + cells_graduated_list[row.Index, 2].Value);
                }
            } 
            #endregion

            #region 驗證原住民名單
            if (wb.Worksheets["原住民名單"] == null)
            {
                MsgBox.Show("Excel檔案 沒有'原住民名單' 頁籤，請確認樣板格式正確", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Worksheet ws_abo_list = wb.Worksheets["原住民名單"];

            Cells cells_abo_list = ws_abo_list.Cells;

            if ("" + cells_abo_list[0, 0].Value != "學號")
            {
                errorList.Add("頁籤:原住民名單，第一行表頭應為:學號");
            }
            if ("" + cells_abo_list[0, 1].Value != "姓名")
            {
                errorList.Add("頁籤:原住民名單，第二行表頭應為:姓名");
            }
            if ("" + cells_abo_list[0, 2].Value != "族別")
            {
                errorList.Add("頁籤:原住民名單，第二行表頭應為:族別");
            }

            // 驗證使用者輸入 原住民名單 是否正確
            foreach (var row in cells_abo_list.Rows)
            {
                if (row.Index > 0 && !row.IsBlank)
                {
                    //Key: 學號+ _ + 姓名
                    if (!stuNumber_Name_list.Contains("" + cells_abo_list[row.Index, 0].Value + "_" + cells_abo_list[row.Index, 1].Value))
                    {
                        errorList.Add("頁籤:原住民名單，第" + (row.Index + 1) + "列，學生:" + cells_abo_list[row.Index, 1].Value + " ，其學號+姓名不存在於畢業生榜單上，請確認輸入格式內容正確。");
                    }
                }
            } 
            #endregion

            #endregion

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

            // 暫停UI 功能
            textBoxX1.Enabled = false;
            buttonX1.Enabled = false;
            buttonX2.Enabled = false;
            buttonX3.Enabled = false;

            //加入背景執行序
            _bw.DoWork += new DoWorkEventHandler(_bkWork_DoWork);

            _bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_worker_RunWorkerCompleted);

            _bw.ProgressChanged += new ProgressChangedEventHandler(_worker_ProgressChanged);

            _bw.WorkerReportsProgress = true;

            _bw.RunWorkerAsync();

        }

        private void _bkWork_DoWork(object sender, DoWorkEventArgs e)
        {
            //載入匯入檔案
            Workbook wb = new Workbook(source_data);

            //  複製高級中等學校應屆畢業生升學就業概況調查表 、高級中等學校原住民應屆畢業生升學就業概況調查表 使用
            Workbook wb_template = new Workbook(new MemoryStream(Properties.Resources.畢業生進路調查公務統計報表));

            Worksheet ws_totalstudent_list = wb_template.Worksheets["高級中等學校應屆畢業生升學就業概況調查表"];
            Worksheet ws_abostudent_list = wb_template.Worksheets["高級中等學校原住民應屆畢業生升學就業概況調查表"];


            wb.Worksheets.Add("高級中等學校應屆畢業生升學就業概況調查表");
            wb.Worksheets.Add("高級中等學校原住民應屆畢業生升學就業概況調查表");

            // 自樣板複製過去
            wb.Worksheets["高級中等學校應屆畢業生升學就業概況調查表"].Copy(ws_totalstudent_list);
            wb.Worksheets["高級中等學校原住民應屆畢業生升學就業概況調查表"].Copy(ws_abostudent_list);

            Worksheet ws_graduated_list = wb.Worksheets["畢業生榜單"];

            Cells cells_graduated_list = ws_graduated_list.Cells;

            Worksheet ws_abo_list = wb.Worksheets["原住民名單"];

            Cells cells_abo_list = ws_abo_list.Cells;
            
            //原住民學生學號 清單
            List<string> abo_number_list = new List<string>();

            foreach (var row in cells_abo_list.Rows)
            {
                if (row.Index > 0 && !row.IsBlank)
                {
                    // 加入原住民學生學號
                    if (!abo_number_list.Contains("" + cells_abo_list[row.Index, 0].Value))
                    {
                        abo_number_list.Add("" + cells_abo_list[row.Index, 0].Value);
                    }
                }
            }

            // 整理所有學生各科別、性別、分類的Dict， Key: 科別_性別_分類代碼， EX: 普通科_男_03私立大學校院日間部(含第二部、四技)，
            // (注意，這邊為了與最後頁籤對應方便，將分類代碼的字首英文字都去掉，A03私立大學校院日間部(含第二部、四技)>> 03私立大學校院日間部(含第二部、四技))
            Dictionary<string, int> arrange_graduated_dict = new Dictionary<string, int>();

            // 整理所有原住民學生各科別、性別、分類的Dict， Key: 科別_性別_分類代碼， EX: 普通科_男_03私立大學校院日間部(含第二部、四技)，
            // (注意，這邊為了與最後頁籤對應方便，將分類代碼的字首英文字都去掉，A03私立大學校院日間部(含第二部、四技)>> 03私立大學校院日間部(含第二部、四技))
            Dictionary<string, int> arrange_graduated_dict_abo = new Dictionary<string, int>();

            //蒐集所有學生所有科別的list
            List<string> dept_list = new List<string>();

            //蒐集原住民學生所有科別的list
            List<string> dept_list_abo = new List<string>();

            int success_count = 0;

            foreach (var row in cells_graduated_list.Rows)
            {
                #region 全校學生
                // 全校學生
                if (row.Index > 0 && !row.IsBlank)
                {
                    // 假如有人工指定分類 以人工指定分類為優先
                    if ("" + cells_graduated_list[row.Index, 9].Value != "")
                    {
                        //假如目前沒有該科別
                        if (!arrange_graduated_dict.ContainsKey("" + cells_graduated_list[row.Index, 1].Value + "_" + cells_graduated_list[row.Index, 3].Value + "_" + ("" + cells_graduated_list[row.Index, 9].Value).Substring(1)))
                        {
                            arrange_graduated_dict.Add("" + cells_graduated_list[row.Index, 1].Value + "_" + cells_graduated_list[row.Index, 3].Value + "_" + ("" + cells_graduated_list[row.Index, 9].Value).Substring(1), 0);

                            arrange_graduated_dict["" + cells_graduated_list[row.Index, 1].Value + "_" + cells_graduated_list[row.Index, 3].Value + "_" + ("" + cells_graduated_list[row.Index, 9].Value).Substring(1)]++;
                        }
                        else
                        {
                            arrange_graduated_dict["" + cells_graduated_list[row.Index, 1].Value + "_" + cells_graduated_list[row.Index, 3].Value + "_" + ("" + cells_graduated_list[row.Index, 9].Value).Substring(1)]++;
                        }
                    }
                    else
                    {
                        //假如目前沒有該科別
                        if (!arrange_graduated_dict.ContainsKey("" + cells_graduated_list[row.Index, 1].Value + "_" + cells_graduated_list[row.Index, 3].Value + "_" + ("" + cells_graduated_list[row.Index, 8].Value).Substring(1)))
                        {
                            arrange_graduated_dict.Add("" + cells_graduated_list[row.Index, 1].Value + "_" + cells_graduated_list[row.Index, 3].Value + "_" + ("" + cells_graduated_list[row.Index, 8].Value).Substring(1), 0);

                            arrange_graduated_dict["" + cells_graduated_list[row.Index, 1].Value + "_" + cells_graduated_list[row.Index, 3].Value + "_" + ("" + cells_graduated_list[row.Index, 8].Value).Substring(1)]++;
                        }
                        else
                        {
                            arrange_graduated_dict["" + cells_graduated_list[row.Index, 1].Value + "_" + cells_graduated_list[row.Index, 3].Value + "_" + ("" + cells_graduated_list[row.Index, 8].Value).Substring(1)]++;
                        }
                    }

                    // 加入科別
                    if (!dept_list.Contains("" + cells_graduated_list[row.Index, 1].Value))
                    {
                        dept_list.Add("" + cells_graduated_list[row.Index, 1].Value);
                    }
                } 
                #endregion

                #region 原住民學生
                //原住民學生
                if (row.Index > 0 && !row.IsBlank && abo_number_list.Contains("" + cells_graduated_list[row.Index, 0].Value))
                {
                    // 假如有人工指定分類 以人工指定分類為優先
                    if ("" + cells_graduated_list[row.Index, 9].Value != "")
                    {
                        //假如目前沒有該科別
                        if (!arrange_graduated_dict_abo.ContainsKey("" + cells_graduated_list[row.Index, 1].Value + "_" + cells_graduated_list[row.Index, 3].Value + "_" + ("" + cells_graduated_list[row.Index, 9].Value).Substring(1)))
                        {
                            arrange_graduated_dict_abo.Add("" + cells_graduated_list[row.Index, 1].Value + "_" + cells_graduated_list[row.Index, 3].Value + "_" + ("" + cells_graduated_list[row.Index, 9].Value).Substring(1), 0);

                            arrange_graduated_dict_abo["" + cells_graduated_list[row.Index, 1].Value + "_" + cells_graduated_list[row.Index, 3].Value + "_" + ("" + cells_graduated_list[row.Index, 9].Value).Substring(1)]++;
                        }
                        else
                        {
                            arrange_graduated_dict_abo["" + cells_graduated_list[row.Index, 1].Value + "_" + cells_graduated_list[row.Index, 3].Value + "_" + ("" + cells_graduated_list[row.Index, 9].Value).Substring(1)]++;
                        }
                    }
                    else
                    {
                        //假如目前沒有該科別
                        if (!arrange_graduated_dict_abo.ContainsKey("" + cells_graduated_list[row.Index, 1].Value + "_" + cells_graduated_list[row.Index, 3].Value + "_" + ("" + cells_graduated_list[row.Index, 8].Value).Substring(1)))
                        {
                            arrange_graduated_dict_abo.Add("" + cells_graduated_list[row.Index, 1].Value + "_" + cells_graduated_list[row.Index, 3].Value + "_" + ("" + cells_graduated_list[row.Index, 8].Value).Substring(1), 0);

                            arrange_graduated_dict_abo["" + cells_graduated_list[row.Index, 1].Value + "_" + cells_graduated_list[row.Index, 3].Value + "_" + ("" + cells_graduated_list[row.Index, 8].Value).Substring(1)]++;
                        }
                        else
                        {
                            arrange_graduated_dict_abo["" + cells_graduated_list[row.Index, 1].Value + "_" + cells_graduated_list[row.Index, 3].Value + "_" + ("" + cells_graduated_list[row.Index, 8].Value).Substring(1)]++;
                        }
                    }

                    // 加入科別
                    if (!dept_list_abo.Contains("" + cells_graduated_list[row.Index, 1].Value))
                    {
                        dept_list_abo.Add("" + cells_graduated_list[row.Index, 1].Value);
                    }
                } 
                #endregion

                try
                {
                    success_count++;

                    int progress = (success_count * 90 / cells_graduated_list.Rows.Count);

                    _bw.ReportProgress(progress);
                }
                catch (Exception ex)
                {
                    MsgBox.Show(ex.Message, "錯誤", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                }
            }

            #region 頁籤:高級中等學校應屆畢業生升學就業概況調查表  填寫

            Worksheet ws_totalstudent_list2 = wb.Worksheets["高級中等學校應屆畢業生升學就業概況調查表"];

            Cells cells_totalstudent_list2 = ws_totalstudent_list2.Cells;

            int dept_col_index = 5;

            //紀錄各科別填的 起始col 位置，EX: ["普通科",5]
            Dictionary<string, int> dept_location_dict = new Dictionary<string, int>();

            // 填上科別
            foreach (string dept in dept_list)
            {
                cells_totalstudent_list2[10, dept_col_index].Value = dept;

                if (!dept_location_dict.ContainsKey(dept))
                {
                    dept_location_dict.Add(dept, dept_col_index);
                }

                dept_col_index = dept_col_index + 2;
            }
            // 填上個分類人數 至頁籤 高級中等學校應屆畢業生升學就業概況調查表
            foreach (KeyValuePair<string, int> arranged_data in arrange_graduated_dict)
            {
                foreach (var row in cells_totalstudent_list2.Rows)
                {
                    //將填寫邏輯 限制在表單 14 列~ 60 列 之間，避免在不必要的地方填寫。
                    if (row.Index > 13 && row.Index < 60 && !row.IsBlank)
                    {
                        if (arranged_data.Key.Contains("" + cells_totalstudent_list2[row.Index, 0].Value))
                        {
                            foreach (string dept in dept_location_dict.Keys)
                            {
                                if (arranged_data.Key.Contains(dept))
                                {
                                    if (arranged_data.Key.Contains("男"))
                                    {
                                        cells_totalstudent_list2[row.Index, dept_location_dict[dept]].Value = arranged_data.Value;

                                    }
                                    if (arranged_data.Key.Contains("女"))
                                    {
                                        cells_totalstudent_list2[row.Index, dept_location_dict[dept] + 1].Value = arranged_data.Value;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            #endregion


            #region 頁籤:高級中等學校原住民應屆畢業生升學就業概況調查表  填寫
            Worksheet ws_totalstudent_list_abo = wb.Worksheets["高級中等學校原住民應屆畢業生升學就業概況調查表"];

            Cells cells_totalstudent_list_abo = ws_totalstudent_list_abo.Cells;

            int dept_col_index_abo = 5;

            //紀錄原住民生各科別填的 起始col 位置，EX: ["普通科",5]
            Dictionary<string, int> dept_location_dict_abo = new Dictionary<string, int>();

            // 填上科別
            foreach (string dept in dept_list_abo)
            {
                cells_totalstudent_list_abo[10, dept_col_index_abo].Value = dept;

                if (!dept_location_dict_abo.ContainsKey(dept))
                {
                    dept_location_dict_abo.Add(dept, dept_col_index_abo);
                }

                dept_col_index_abo = dept_col_index_abo + 2;
            }
            // 填上個分類人數 至頁籤 高級中等學校原住民應屆畢業生升學就業概況調查表
            foreach (KeyValuePair<string, int> arranged_data in arrange_graduated_dict_abo)
            {
                foreach (var row in cells_totalstudent_list_abo.Rows)
                {
                    //將填寫邏輯 限制在表單 14 列~ 60 列 之間，避免在不必要的地方填寫。
                    if (row.Index > 13 && row.Index < 60 && !row.IsBlank)
                    {
                        if (arranged_data.Key.Contains("" + cells_totalstudent_list_abo[row.Index, 0].Value))
                        {
                            foreach (string dept in dept_location_dict_abo.Keys)
                            {
                                if (arranged_data.Key.Contains(dept))
                                {
                                    if (arranged_data.Key.Contains("男"))
                                    {
                                        cells_totalstudent_list_abo[row.Index, dept_location_dict_abo[dept]].Value = arranged_data.Value;

                                    }
                                    if (arranged_data.Key.Contains("女"))
                                    {
                                        cells_totalstudent_list_abo[row.Index, dept_location_dict_abo[dept] + 1].Value = arranged_data.Value;
                                    }
                                }
                            }
                        }
                    }
                }
            } 
            #endregion


            e.Result = wb;
        }


        private void _worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("處理中...", e.ProgressPercentage);
        }

        private void _worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Workbook wb = (Workbook)e.Result;

            SaveFileDialog sd = new SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = "畢業生進路調查公務統計報表.xlsx";
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

            FISCA.Presentation.MotherForm.SetStatusBarMessage("畢業生進路調查公務統計報表 產生完成");
            // 任務結束，關閉
            this.Close();
        }
        
    }
}
