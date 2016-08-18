using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DevComponents.DotNetBar;
using System.Xml.Linq;

namespace Counsel_System2.Forms
{
    public partial class QuizForm : FISCA.Presentation.Controls.BaseForm
    {
        DAO.UDTTransfer _UDTTransfer;
        DAO.LogTransfer _LogTransfer;
        string _selQuizName = "";
       
        // 測驗資料
        List<DAO.UDT_QuizDef> _QuizData;
        public QuizForm()
        {
            InitializeComponent();
            this.MaximumSize = this.MinimumSize = this.Size;
            _UDTTransfer = new DAO.UDTTransfer();
            _QuizData = new List<DAO.UDT_QuizDef>();
            _LogTransfer = new DAO.LogTransfer();
          
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void QuizForm_Load(object sender, EventArgs e)
        {
            LoadData();
        }
          
        
        private void btnAdd_Click(object sender, EventArgs e)
        {
            List<string> NameList = _QuizData.Select(x => x.QuizName).ToList();

            AddQuizNameForm aqnf = new AddQuizNameForm(NameList);            

            if (aqnf.ShowDialog() == System.Windows.Forms.DialogResult.Yes)
            {
                if (!string.IsNullOrEmpty(aqnf.GetAddQuizName()))
                {
                    DAO.UDT_QuizDef qd = new DAO.UDT_QuizDef();                    
                    qd.QuizName = aqnf.GetAddQuizName();
                    // Log
                    _LogTransfer.Clear();
                    _LogTransfer.SetLogValue("測驗名稱","");
                    _LogTransfer.SetLogValue("測驗名稱", qd.QuizName);
                    _LogTransfer.SaveInsertLog("輔導新增測驗名稱", "新增", "", "", "", "");
                    
                    _UDTTransfer.InsertQuizData(qd);
                    
                }
            }
            LoadData();
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            if (itmPnlQuizName.SelectedItems.Count == 1)
            {
                
                ButtonItem selItm = itmPnlQuizName.SelectedItems[0] as ButtonItem;
                if (selItm != null)
                {                    
                    if (FISCA.Presentation.Controls.MsgBox.Show("請問確定刪除["+selItm.Name+"]?","提示",MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.Yes)
                    {
                        DAO.UDT_QuizDef data = itmPnlQuizName.SelectedItems[0].Tag as DAO.UDT_QuizDef;

                        string itemName = "";
                        List<string> itemList = (from da in Utility.ConvertStringToXelm1(data.QuizDataField).Elements("Field") select da.Attribute("name").Value).ToList();
                        itemName += string.Join(",", itemList.ToArray());
                        _LogTransfer.SetLogValue("測驗名稱", data.QuizName);
                        _LogTransfer.SetLogValue("項目內容", itemName);
                        _LogTransfer.SaveDeleteLog("輔導刪除測驗", "刪除", "", "", "", "");

                        _UDTTransfer.DeleteQuizData(data);
                        LoadData();
                    }
                } 
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            // 檢查資料是否又誤
            if (itmPnlQuizName.SelectedItems.Count < 1)
            {
                FISCA.Presentation.Controls.MsgBox.Show("請選擇測驗.");
                return;
            }

            foreach(DataGridViewRow dr in dgDataField.Rows )
            {
                int errCot = 0;

                if (dr.Cells[colFieldName.Index].ErrorText != "")
                    errCot++;

                if (dr.Cells[colDisplayOrder.Index].ErrorText != "")
                    errCot++;

                if (errCot > 0)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("資料有錯誤無法儲存.");
                    return;                
                }
            }


            // 有修改名稱
            if (txtQuizName.Text.Trim() != _selQuizName)
            {
                if (_QuizData.Select(x=>x.QuizName).Contains(txtQuizName.Text.Trim()))
                {
                    FISCA.Presentation.Controls.MsgBox.Show("已有相同名稱，無法修改!");
                    return;
                }
            }

            // 儲存資料
                if (itmPnlQuizName.SelectedItems.Count == 1)
                {
                    ButtonItem selItem = itmPnlQuizName.SelectedItems[0] as ButtonItem;
                    DAO.UDT_QuizDef qd1 = selItem.Tag as DAO.UDT_QuizDef;

                    if (selItem != null && qd1 != null)
                    {
                        qd1.QuizName = txtQuizName.Text.Trim();

                        List<XElement> xmlList = new List<XElement>();
                        List<string> itemList = new List<string>();
                        foreach (DataGridViewRow dgvr in dgDataField.Rows)
                        {
                            if (dgvr.IsNewRow)
                                continue;
                            
                            XElement elm = new XElement("Field");
                            if (dgvr.Cells[colFieldName.Index].Value != null)
                            {
                                elm.SetAttributeValue("name", dgvr.Cells[colFieldName.Index].Value.ToString());
                                itemList.Add(dgvr.Cells[colFieldName.Index].Value.ToString());
                            }

                            if (dgvr.Cells[colDisplayOrder.Index].Value != null)
                                elm.SetAttributeValue("order", dgvr.Cells[colDisplayOrder.Index].Value.ToString());

                            xmlList.Add(elm);
                        }
                        _LogTransfer.SetLogValue("測驗名稱",txtQuizName.Text);
                        _LogTransfer.SetLogValue("項目內容",string.Join(",",itemList.ToArray()));

                        qd1.QuizDataField = Utility.ConvertXmlListToString1(xmlList);

                        if (string.IsNullOrEmpty(qd1.UID))
                        {                            
                            _LogTransfer.SaveInsertLog("輔導新增測驗", "新增", "", "", "", "");

                            _UDTTransfer.InsertQuizData(qd1);
                        }
                        else
                        {
                            _LogTransfer.SaveInsertLog("輔導修改測驗", "修改", "", "", "", "");
                            _UDTTransfer.UpdateQuizData(qd1);
                        }
                    }
                }
                FISCA.Presentation.Controls.MsgBox.Show("儲存完成.");
            
            LoadData();
        }

        private void LoadData()
        {
            itmPnlQuizName.SuspendLayout();
            itmPnlQuizName.Items.Clear();
            txtQuizName.Text = "";
            dgDataField.Rows.Clear();
            _QuizData=_UDTTransfer.GetAllQuizData();
            _QuizData = (from data in _QuizData orderby data.QuizName select data).ToList();
            
            foreach (DAO.UDT_QuizDef qd in _QuizData)
            {
                ButtonItem btnItem = new ButtonItem();
                btnItem.Name = qd.QuizName;
                btnItem.Text = qd.QuizName;
                btnItem.Tag = qd;
                btnItem.OptionGroup = "itmPnlQuizName";
                btnItem.ButtonStyle = DevComponents.DotNetBar.eButtonStyle.ImageAndText;
                btnItem.Click += new EventHandler(btnItem_Click);
                itmPnlQuizName.Items.Add(btnItem);
            }
           
            
            itmPnlQuizName.ResumeLayout();
            itmPnlQuizName.Refresh();
            txtQuizName.Enabled = false;
        }

        void btnItem_Click(object sender, EventArgs e)
        {
            if (itmPnlQuizName.SelectedItems.Count == 1)
            {
                txtQuizName.Enabled = true;
                ButtonItem selItem = itmPnlQuizName.SelectedItems[0] as ButtonItem;
                if (selItem != null)
                {
                    txtQuizName.Text = selItem.Name;
                    _selQuizName = selItem.Name;
                    dgDataField.Rows.Clear();
                        int RowIdx = 0;
                        DAO.UDT_QuizDef qd1 = selItem.Tag as DAO.UDT_QuizDef;
                        List<string> itemList = new List<string>();
                        if (qd1 != null)
                        {
                            XElement elms = Utility.ConvertStringToXelm1(qd1.QuizDataField);
                            if (elms != null)
                            { 
                                foreach(XElement elm in elms.Elements("Field"))
                                {
                                    RowIdx = dgDataField.Rows.Add();
                                    if (elm.Attribute("name") != null)
                                    {
                                        dgDataField.Rows[RowIdx].Cells[colFieldName.Index].Value = elm.Attribute("name").Value;
                                        itemList.Add(elm.Attribute("name").Value);
                                    }

                                    if (elm.Attribute("order") != null)
                                        dgDataField.Rows[RowIdx].Cells[colDisplayOrder.Index].Value = elm.Attribute("order").Value;
                                }                            
                            }                        
                        }

                        _LogTransfer.Clear();
                        _LogTransfer.SetLogValue("測驗名稱", qd1.QuizName);
                        _LogTransfer.SetLogValue("項目內容", string.Join(",",itemList.ToArray()));
                }
                else
                {
                    txtQuizName.Text = "";
                    dgDataField.Rows.Clear();
                }

                
            }
        }

        private void dgDataField_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            dgDataField.EndEdit();

            if (dgDataField.CurrentCell.ColumnIndex == colDisplayOrder.Index)
            {
                if (dgDataField.CurrentCell.Value != null)
                {
                    int d;
                    if (int.TryParse(dgDataField.CurrentCell.Value.ToString(), out d))
                        dgDataField.CurrentCell.ErrorText = "";
                    else
                        dgDataField.CurrentCell.ErrorText = "必須整數";
                }
            }

            if (dgDataField.CurrentCell.ColumnIndex == colFieldName.Index )
            {
                int co = 0;

                if (dgDataField.CurrentCell.Value != null)
                {
                    foreach (DataGridViewRow dr in dgDataField.Rows)
                    {
                        if (dr.IsNewRow)
                            continue;
                        if (dr.Cells[colFieldName.Index].Value != null)
                            if (dr.Cells[colFieldName.Index].Value.ToString().Trim() == dgDataField.CurrentCell.Value.ToString().Trim())
                                co++;
                    }                    
                    dgDataField.CurrentCell.ErrorText = "";
                    if(co>1)
                        dgDataField.CurrentCell.ErrorText = "欄位名稱不能重複!";
                }
            }


            dgDataField.BeginEdit(false);
        }

        private void 新增ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            List<string> NameList = _QuizData.Select(x => x.QuizName).ToList();

            AddQuizNameForm aqnf = new AddQuizNameForm(NameList);

            if (aqnf.ShowDialog() == System.Windows.Forms.DialogResult.Yes)
            {
                if (!string.IsNullOrEmpty(aqnf.GetAddQuizName()))
                {
                    DAO.UDT_QuizDef qd = new DAO.UDT_QuizDef();
                    qd.QuizName = aqnf.GetAddQuizName();
                    // Log
                    _LogTransfer.Clear();
                    _LogTransfer.SetLogValue("測驗名稱", "");
                    _LogTransfer.SetLogValue("測驗名稱", qd.QuizName);
                    _LogTransfer.SaveInsertLog("輔導新增測驗名稱", "新增", "", "", "", "");

                    _UDTTransfer.InsertQuizData(qd);

                }
            }
            LoadData();
        }

        private void 刪除ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (itmPnlQuizName.SelectedItems.Count == 1)
            {

                ButtonItem selItm = itmPnlQuizName.SelectedItems[0] as ButtonItem;
                if (selItm != null)
                {
                    if (FISCA.Presentation.Controls.MsgBox.Show("請問確定刪除[" + selItm.Name + "]?", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.Yes)
                    {
                        DAO.UDT_QuizDef data = itmPnlQuizName.SelectedItems[0].Tag as DAO.UDT_QuizDef;

                        string itemName = "";
                        List<string> itemList = (from da in Utility.ConvertStringToXelm1(data.QuizDataField).Elements("Field") select da.Attribute("name").Value).ToList();
                        itemName += string.Join(",", itemList.ToArray());
                        _LogTransfer.SetLogValue("測驗名稱", data.QuizName);
                        _LogTransfer.SetLogValue("項目內容", itemName);
                        _LogTransfer.SaveDeleteLog("輔導刪除測驗", "刪除", "", "", "", "");

                        _UDTTransfer.DeleteQuizData(data);
                        LoadData();
                    }
                }
            }
        }
    }
}
