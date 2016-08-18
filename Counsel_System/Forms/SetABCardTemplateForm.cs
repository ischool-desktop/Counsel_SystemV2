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
    /// <summary>
    /// 設定綜合表現樣版
    /// </summary>
    public partial class SetABCardTemplateForm : FISCA.Presentation.Controls.BaseForm 
    {

        /// <summary>
        /// Log 處理
        /// </summary>
        DAO.LogTransfer _LogTransfer;
        /// <summary>
        /// 樣板交換處理
        /// </summary>
        DAO.ABCardTemplateTransfer _ABCardTemplateTransfer;
        
        BackgroundWorker _bgWorker;
        bool isBGBusy = false;

        public SetABCardTemplateForm()
        {
            InitializeComponent();
            _LogTransfer = new DAO.LogTransfer();            
            _ABCardTemplateTransfer = new DAO.ABCardTemplateTransfer();
            this.MaximumSize = this.MinimumSize = this.Size;
            _bgWorker = new BackgroundWorker();
            _bgWorker.DoWork += new DoWorkEventHandler(_bgWorker_DoWork);
            _bgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgWorker_RunWorkerCompleted);
            dgDataField.DataError += new DataGridViewDataErrorEventHandler(dgDataField_DataError);
            colQqType.Items.AddRange(DAO.ABCardTemplateTransfer.GetQTypeList().ToArray());
            
        }

        void dgDataField_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            
        }

        void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            if (isBGBusy)
            {
                isBGBusy = false;
                _bgWorker.RunWorkerAsync();
                return;
            }
            LoadData();
    
        }

        private void _BGRun()
        {
            if (_bgWorker.IsBusy)
                isBGBusy = true;
            else
                _bgWorker.RunWorkerAsync();
        }

        void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _ABCardTemplateTransfer.LoadAllTemplate(DAO.UDTTransfer.GetABCardTemplate());
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            // 檢查畫面是否選擇
            if (itmPnlSubjectName.SelectedItems.Count == 1)
            {
                
                // 讀取畫面上設定與相對欄位資料
                ButtonItem selItem = itmPnlSubjectName.SelectedItems[0] as ButtonItem;
                //if (selItem != null && _ABCardTemplateDefinitionDefDict.ContainsKey(selItem.Name))
                //{   
                //    // 透過 UDT 更新資料



                //    // 重新載入畫面
                //    LoadData();
                //}
                //else
                //    FISCA.Presentation.Controls.MsgBox.Show("讀取資料失敗.");
            }
            else
                FISCA.Presentation.Controls.MsgBox.Show("請選擇資料.");
        }

        private void LoadData()
        {   
            // 載入畫面
            itmPnlSubjectName.SuspendLayout();
            itmPnlSubjectName.Items.Clear();
            txtSubjectName.Text = "";            
            dgDataField.Rows.Clear();


                foreach (string name in _ABCardTemplateTransfer.GetAllSubjectNameList())
                {
                    ButtonItem btnItem = new ButtonItem();
                    btnItem.Name = name;
                    btnItem.Text = name;                    
                    btnItem.OptionGroup = "itmPnlSubjectName";
                    btnItem.ButtonStyle = DevComponents.DotNetBar.eButtonStyle.ImageAndText;
                    btnItem.Click += new EventHandler(btnItem_Click);
                    itmPnlSubjectName.Items.Add(btnItem);
                }

            itmPnlSubjectName.ResumeLayout();
            itmPnlSubjectName.Refresh();        
        }

        void btnItem_Click(object sender, EventArgs e)
        {

            if (itmPnlSubjectName.SelectedItems.Count == 1)
            {                
                // 將所選的資料載入右邊資料向內
                ButtonItem selItem = itmPnlSubjectName.SelectedItems[0] as ButtonItem;
                if (selItem != null)
                {
                    txtSubjectName.Text = selItem.Text;
                    dgDataField.Rows.Clear();
                    foreach (DAO.QuestionPKey qpk in _ABCardTemplateTransfer.GetQuestionListBySubjectLabel(txtSubjectName.Text))
                    {
                        int row = dgDataField.Rows.Add();
                        dgDataField.Rows[row].Tag = qpk;
                        dgDataField.Rows[row].Cells[colQGLabel.Index].Value = qpk.QGLabel;
                        dgDataField.Rows[row].Cells[colQqLabel.Index].Value = qpk.QLabel;
                        dgDataField.Rows[row].Cells[colQqType.Index].Value = qpk.QType;
                        if(qpk.dataElement ==null)
                            dgDataField.Rows[row].Cells[colQqDetail.Index].Value = "   ";
                        else
                            dgDataField.Rows[row].Cells[colQqDetail.Index].Value = "...";
                    }
                }
                else
                {
                    txtSubjectName.Text = "";
                    dgDataField.Rows.Clear();
                }
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            // 呼叫新增畫面
            Forms.AddABCardSubjectNameForm asnf = new AddABCardSubjectNameForm(_ABCardTemplateTransfer.GetAllSubjectNameList());

            if (asnf.ShowDialog() == System.Windows.Forms.DialogResult.Yes)
            {
                string SubjectName = asnf.GetAddSubjectName();
                if (!string.IsNullOrEmpty(SubjectName))
                {
                    XElement elm = new XElement("Subject");
                    elm.SetAttributeValue("label", SubjectName);
                    DAO.UDT_ABCardTemplateDefinitionDef abtdf = new DAO.UDT_ABCardTemplateDefinitionDef();
                    abtdf.SubjectName = SubjectName;
                    abtdf.Content = elm.ToString();

                    // 呼叫UDT 新增資料
                    List<DAO.UDT_ABCardTemplateDefinitionDef> addList = new List<DAO.UDT_ABCardTemplateDefinitionDef>();
                    addList.Add(abtdf);
                    DAO.UDTTransfer.InsertABCardTemplate(addList);

                    // 畫面重新載入
                    _BGRun();
                }
            }
        }

        
        private void SetABCardTemplateForm_Load(object sender, EventArgs e)
        {
            // 讀取資料                   
            _BGRun();
 
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            // 檢查畫面是否有選
            if (itmPnlSubjectName.SelectedItems.Count == 1)
            {
                ButtonItem selItem = itmPnlSubjectName.SelectedItems[0] as ButtonItem;
                if (selItem != null)
                {
                    // 提示是否刪除
                    if (FISCA.Presentation.Controls.MsgBox.Show("請問確認刪除[" + selItem.Name + "]?", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.Yes)
                    {
                        // 收集刪除資料
                        List<DAO.UDT_ABCardTemplateDefinitionDef> delList = new List<DAO.UDT_ABCardTemplateDefinitionDef> ();
                        DAO.UDT_ABCardTemplateDefinitionDef abtd = _ABCardTemplateTransfer.GetUDTTemplateBySubjectLabel(selItem.Name);

                        if(abtd !=null)
                            delList.Add(abtd);

                        // 執行刪除
                        DAO.UDTTransfer.DeleteABCardTemplate(delList);

                        // 重新載入畫面
                        _BGRun();
                    }
                }
            }
            else
                FISCA.Presentation.Controls.MsgBox.Show("請選擇資料");
        }

        private void dgDataField_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == colQqDetail.Index)
            {
                string type = dgDataField.Rows[e.RowIndex].Cells[colQqType.Index].Value.ToString();

                if (type == "text" || type == "textarea")
                { 
                
                }

                if (type == "combobox" || type == "textboxdropdown" || type == "checkbox")
                { 
                
                
                }

                if (type == "grid")
                { 
                
                }
                
            }
        }
    }
}
