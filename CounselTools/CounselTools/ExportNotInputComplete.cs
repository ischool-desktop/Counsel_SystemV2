using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Cells;
using System.ComponentModel;
using System.Text;
using System.Data;
using System.IO;

namespace CounselTools
{
    public class ExportNotInputABCardComplete
    {
        List<string> _StudentIDList;
        BackgroundWorker _bgLoadData;

        

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
            FISCA.Presentation.MotherForm.SetStatusBarMessage("綜合紀錄表未輸入完整名單產生中", e.ProgressPercentage);
        }

        void _bgLoadData_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error !=null)
            {
                FISCA.Presentation.Controls.MsgBox.Show("產生過程發生錯誤,"+e.Error.Message);
            }
            else
            {
                try
                {
                    Workbook wb = e.Result as Workbook;
                    if (wb != null)
                    {
                        Utility.CompletedXls("輔導綜合紀錄表未輸入完整名單", wb);
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
                cpGNameList.Add("本人概況");
                cpGNameList.Add("家庭狀況");
                cpGNameList.Add("學習狀況");
                cpGNameList.Add("自傳");
                cpGNameList.Add("自我認識");
                cpGNameList.Add("生活感想");
                cpGNameList.Add("畢業後計畫");
                cpGNameList.Add("備註");
                cpGNameList.Add("適應情形");

                foreach (ClassStudent cs in ClassStudents)
                {
                    Dictionary<string, ICheckProcess> CheckProcDict = new Dictionary<string, ICheckProcess>();        
                    // 開始檢查
                    foreach (string cpGName in cpGNameList)
                    {
                        switch (cpGName)
                        { 
                            case "本人概況":
                                CheckProcDict.Add(cpGName, new CheckProcess1());
                                break;
                            case "家庭狀況":
                                CheckProcDict.Add(cpGName, new CheckProcess2());
                                break;
                            case "學習狀況":
                                CheckProcDict.Add(cpGName, new CheckProcess3());
                                break;
                            case "自傳":
                                CheckProcDict.Add(cpGName, new CheckProcess4());
                                break;
                            case "自我認識":
                                CheckProcDict.Add(cpGName, new CheckProcess5());
                                break;
                            case "生活感想":
                                CheckProcDict.Add(cpGName, new CheckProcess6());
                                break;
                            case "畢業後計畫":
                                CheckProcDict.Add(cpGName, new CheckProcess7());
                                break;
                            case "備註":
                                CheckProcDict.Add(cpGName, new CheckProcess8());
                                break;
                            case "適應情形":
                                CheckProcDict.Add(cpGName, new CheckProcess9());
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
                            if (cp.GetErrorCount() > 0)
                            {
                                if (!cs.NonInputCompleteDict.ContainsKey(cpGName))
                                    cs.NonInputCompleteDict.Add(cpGName, cp.GetMessage());
                            }
                        }
                    }
                }

                _bgLoadData.ReportProgress(80);

                // 讀取樣版
                Workbook wb = new Workbook(new MemoryStream(Properties.Resources.未輸入完整名單樣版));

                // 綜合紀錄索引
                Dictionary<string, int> gpColIdx = new Dictionary<string, int>();
                int col = 5;
                foreach (string cpName in cpGNameList)
                    gpColIdx.Add(cpName, col++);

                // 學號,班級,座號,姓名,本人概況,家庭狀況,學習狀況,自傳,自我認識,生活感想,畢業後計畫,備註,適應情形
                int rowIdx = 1;
                foreach (ClassStudent cs in ClassStudents)
                {
                    // 有缺才填入
                    if (cs.NonInputCompleteDict.Count > 0)
                    {
                        wb.Worksheets[0].Cells[rowIdx, 0].PutValue(cs.StudentNumber);
                        wb.Worksheets[0].Cells[rowIdx, 1].PutValue(cs.GradeYearDisplay);
                        wb.Worksheets[0].Cells[rowIdx,2].PutValue(cs.ClassName);
                        wb.Worksheets[0].Cells[rowIdx, 3].PutValue(cs.SeatNo);
                        wb.Worksheets[0].Cells[rowIdx, 4].PutValue(cs.StudentName);
                        // 填入缺漏資料
                        foreach (string key in cs.NonInputCompleteDict.Keys)
                        {
                            if (gpColIdx.ContainsKey(key))
                                wb.Worksheets[0].Cells[rowIdx, gpColIdx[key]].PutValue(cs.NonInputCompleteDict[key]);
                        }
                        rowIdx++;
                    }                
                }
                _bgLoadData.ReportProgress(95);

                wb.Worksheets[0].AutoFitColumns();
                e.Result = wb;

                _bgLoadData.ReportProgress(100);
            }
            catch(Exception ex) {


                e.Cancel = true;
            }
        }
               
    }
}
