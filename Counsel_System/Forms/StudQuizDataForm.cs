using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using K12.Data;
using System.Xml.Linq;

namespace Counsel_System2.Forms
{
    public partial class StudQuizDataForm : FISCA.Presentation.Controls.BaseForm
    {
        public enum EditMode {Insert,Edit}
        DAO.UDT_StudQuizDataDef _StudQuizData;
        DAO.UDTTransfer _UDTTransfer;
        DAO.UDT_QuizDef _qdd;
        List<DAO.SortData1> _quizFieldNameList;
        private EditMode _EditMode;
        // 所有測驗
        List<DAO.UDT_QuizDef> _AllQuiz;
        // 學生測驗資料檢查用
        List<DAO.UDT_StudQuizDataDef> _StudQuizDataList;
        List<string> _RowNameList;
        string _StudentID;
        DAO.LogTransfer _LogTransfer;
        string _LogStudentName = "";

        public StudQuizDataForm(DAO.UDT_StudQuizDataDef sqd,EditMode editMode,string StudentID)
        {
            InitializeComponent();
            _EditMode = editMode;
            _RowNameList = new List<string>();
            _StudQuizData = sqd;
            _UDTTransfer = new DAO.UDTTransfer();
            _AllQuiz = _UDTTransfer.GetAllQuizData();
            _quizFieldNameList = new List<DAO.SortData1>();
            _StudentID = StudentID;
            _LogTransfer = new DAO.LogTransfer();            

            // 將測驗名稱放入
            List<string> nameList = (from data in _AllQuiz orderby data.QuizName select data.QuizName).ToList();
            cbxQuizName.Items.AddRange(nameList.ToArray());
            
            ReLoadQuizFieldNameList();
           
        }

        private void LogData()
        {
            _LogTransfer.SetLogValue("測驗名稱",cbxQuizName.Text);
            if (dtImplementationDate.IsEmpty)
                _LogTransfer.SetLogValue("實施日期", "");
            else
                _LogTransfer.SetLogValue("實施日期", dtImplementationDate.Value.ToShortDateString());

            if (dtAnalysisDate.IsEmpty)
                _LogTransfer.SetLogValue("解析日期", "");
            else
                _LogTransfer.SetLogValue("解析日期", dtAnalysisDate.Value.ToShortDateString());

            foreach (DataGridViewRow drv in dgQuizData.Rows)
            {
                if (drv.IsNewRow)
                    continue;
                if (drv.Cells[colDataField.Index].Value!=null)
                {
                    if(drv.Cells[colDataValue.Index].Value == null)
                        _LogTransfer.SetLogValue("項目名稱："+drv.Cells[colDataField.Index].Value.ToString(), "");
                    else                
                        _LogTransfer.SetLogValue("項目名稱："+drv.Cells[colDataField.Index].Value.ToString(), "測驗結果："+drv.Cells[colDataValue.Index].Value.ToString());            
                }
            }
        }

        private void ReLoadQuizFieldNameList()
        {
            if (_qdd != null)
            {
                _quizFieldNameList.Clear();
                XElement elms = Utility.ConvertStringToXelm1(_qdd.QuizDataField);
                foreach (XElement elm in elms.Elements("Field"))
                {
                    DAO.SortData1 sd = new DAO.SortData1();
                    sd.Name = elm.Attribute("name").Value;
                    if (elm.Attribute("order") == null)
                        sd.Order = 999;
                    else
                        sd.Order = int.Parse(elm.Attribute("order").Value);

                    _quizFieldNameList.Add(sd);
                }
                _quizFieldNameList = (from data in _quizFieldNameList orderby data.Order select data).ToList();
                _RowNameList = (from data in _quizFieldNameList select data.Name).ToList();
            }       
        }


        private void StudQuizDataForm_Load(object sender, EventArgs e)
        {
            // 取得學生資料
            StudentRecord studRec = Student.SelectByID(_StudentID);
            _StudQuizDataList = _UDTTransfer.GetStudQuizDataByStudentID(studRec.ID);
            
            // log 學生資訊用
            _LogStudentName = Utility.ConvertString1(studRec);


            if (_StudQuizData == null)
            {
                _StudQuizData = new DAO.UDT_StudQuizDataDef();
                
            }

            if (_EditMode == EditMode.Edit)
            {
                // 取得測驗名稱
                foreach (DAO.UDT_QuizDef data in _AllQuiz.Where(x => x.UID == _StudQuizData.QuizID.ToString()))
                    _qdd = data;

                if (_qdd == null)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("無法解析測驗內容.");
                    return;
                }
            }

            if (studRec != null)
            {
                string str = "";
                if(studRec.Class !=null )
                    str+=studRec.Class.Name+" ";

                str+=studRec.Name+" ";
                str += studRec.StudentNumber;
                str = " (" + str + ")";
                if (_EditMode == EditMode.Edit)
                    this.Text = _qdd.QuizName + str;
                else
                    this.Text = str;
            }
            
            // 當編輯模式，測驗名稱,不能修改
            if (_EditMode == EditMode.Edit)
            {
                cbxQuizName.Text = _qdd.QuizName;
                cbxQuizName.Enabled = false;

                if (_StudQuizData.ImplementationDate.HasValue)
                    dtImplementationDate.Value = _StudQuizData.ImplementationDate.Value;

                if (_StudQuizData.AnalysisDate.HasValue)
                    dtAnalysisDate.Value = _StudQuizData.AnalysisDate.Value;

                dgQuizData.Rows.Clear();

                int rowIdx = 0;
                dgQuizData.Columns[colDataField.Index].ReadOnly = false;

                XElement elmContent = Utility.ConvertStringToXelm1(_StudQuizData.Content);
                if (elmContent != null)
                {
                    // 沒有在試別欄位加入
                    List<string> fieldList = (from data in elmContent.Elements("Item") select data.Attribute("name").Value).ToList();
                    foreach (string str in fieldList)
                        if (!_RowNameList.Contains(str))
                            _RowNameList.Add(str);
                    // 依試別欄位順序排序
                    foreach (string str in _RowNameList)
                    {
                        foreach (XElement elm in elmContent.Elements("Item").Where(x => x.Attribute("name").Value == str))
                        {
                            rowIdx = dgQuizData.Rows.Add();
                            if (elm.Attribute("name") != null)
                                dgQuizData.Rows[rowIdx].Cells[colDataField.Index].Value = elm.Attribute("name").Value;
                            if (elm.Attribute("value") != null)
                                dgQuizData.Rows[rowIdx].Cells[colDataValue.Index].Value = elm.Attribute("value").Value;
                        }
                    }
                }
                dgQuizData.Columns[colDataField.Index].ReadOnly = true;
            }
            _LogTransfer.Clear();
            LogData();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(cbxQuizName.Text))
            {
                FISCA.Presentation.Controls.MsgBox.Show("請選擇測驗名稱.");
                return;
            }

            if (dtAnalysisDate.IsEmpty)
            {
                FISCA.Presentation.Controls.MsgBox.Show("請輸入解析日期.");
                return;
            }

            if (dtImplementationDate.IsEmpty)
            {
                FISCA.Presentation.Controls.MsgBox.Show("請輸入實施日期.");
                return;
            }

            foreach (DataGridViewRow dr in dgQuizData.Rows)
            {
                int errCot = 0;

                if (dr.Cells[colDataField.Index].ErrorText != "")
                    errCot++;

                if (errCot > 0)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("資料有錯誤無法儲存.");
                    return;
                }
            }

            LogData();

            if (_StudQuizData == null)
                _StudQuizData = new DAO.UDT_StudQuizDataDef();

            _StudQuizData.StudentID = int.Parse(_StudentID);
            
            _StudQuizData.ImplementationDate = dtImplementationDate.Value;
            _StudQuizData.AnalysisDate = dtAnalysisDate.Value;

            List<XElement> xmlList = new List<XElement>();
            foreach (DataGridViewRow dgvr in dgQuizData.Rows)
            {
                if (dgvr.IsNewRow)
                    continue;
                XElement elm = new XElement("Item");
                
                if(dgvr.Cells[colDataField.Index].Value !=null )
                    elm.SetAttributeValue("name", dgvr.Cells[colDataField.Index].Value.ToString());
                if (dgvr.Cells[colDataValue.Index].Value != null)
                    elm.SetAttributeValue("value", dgvr.Cells[colDataValue.Index].Value.ToString());
                xmlList.Add(elm);
            }
            _StudQuizData.Content = Utility.ConvertXmlListToString1(xmlList);
            if (_EditMode == EditMode.Edit)
            {
                _UDTTransfer.UpdateStudQuizData(_StudQuizData);
                // log
                _LogTransfer.SaveChangeLog("學生.輔導測驗相關-修改", "修改",_LogStudentName+",修改測驗名稱："+cbxQuizName.Text+"\n", "", "student",_StudentID);       
            }
            else
            {
                // 取得所選測驗ID
                foreach (DAO.UDT_QuizDef data in _AllQuiz.Where(x => x.QuizName == cbxQuizName.Text))
                    _StudQuizData.QuizID = int.Parse(data.UID);


                // 檢查學生測驗是否有名稱id
                if (_StudQuizData.QuizID < 1)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("測驗名稱無法對應.");
                    return;
                }

                bool pass = true;
                // 檢查學生測驗資料是否有相同:測驗名稱
                if (_StudQuizDataList == null)
                    _StudQuizDataList = new List<DAO.UDT_StudQuizDataDef>();

                foreach (DAO.UDT_StudQuizDataDef data in _StudQuizDataList.Where(x => x.QuizID == _StudQuizData.QuizID))
                        pass = false;

                if (pass == false)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("已有相同測驗名稱無法新增.");
                    return;
                }

                if (pass)
                {

                    _UDTTransfer.InsertStudQuizData(_StudQuizData);
                    _LogTransfer.SaveInsertLog("學生.輔導相關測驗-新增", "新增",_LogStudentName+",新增測驗名稱："+cbxQuizName.Text+"\n","", "student", _StudentID);
                }
            }
           
            _LogTransfer.Clear();
            LogData();
            this.Close();
        }

        private void dgQuizData_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            dgQuizData.EndEdit();

            if (dgQuizData.CurrentCell.ColumnIndex == colDataField.Index)
            {
                int co = 0;

                if (dgQuizData.CurrentCell.Value != null)
                {
                    foreach (DataGridViewRow dr in dgQuizData.Rows)
                    {
                        if (dr.IsNewRow)
                            continue;
                        if (dr.Cells[colDataField.Index].Value != null)
                            if (dr.Cells[colDataField.Index].Value.ToString().Trim() == dgQuizData.CurrentCell.Value.ToString().Trim())
                                co++;
                    }
                    dgQuizData.CurrentCell.ErrorText = "";
                    if (co > 1)
                        dgQuizData.CurrentCell.ErrorText = "欄位名稱不能重複!";
                }
            }
            dgQuizData.BeginEdit(false);
        }

        private void cbxQuizName_SelectedIndexChanged(object sender, EventArgs e)
        {
            foreach (DAO.UDT_QuizDef data in _AllQuiz.Where(x => x.QuizName == cbxQuizName.Text))
                _qdd = data;

            ReLoadQuizFieldNameList();

            dgQuizData.Rows.Clear();

            int rowIdx = 0;
            dgQuizData.Columns[colDataField.Index].ReadOnly = false;
         
                // 依試別欄位順序排序
                foreach (DAO.SortData1 data in _quizFieldNameList)
                {
                    rowIdx = dgQuizData.Rows.Add();
                    dgQuizData.Rows[rowIdx].Cells[colDataField.Index].Value = data.Name;
                }
         
            dgQuizData.Columns[colDataField.Index].ReadOnly = true;
            _LogTransfer.Clear();
            LogData();
        }        
    }
}
