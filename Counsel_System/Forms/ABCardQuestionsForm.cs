using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Counsel_System2.DAO;
using DevComponents.DotNetBar;


namespace Counsel_System2.Forms
{
    public partial class ABCardQuestionsForm : FISCA.Presentation.Controls.BaseForm
    {
        List<string> _GroupList;
        QuestionData _QuestionData = null;
        ABCardQuestionDataManager _ABCardQuestionDataManager;
        Dictionary<string, EnumControlType> _ControlTypeDictE = new Dictionary<string, EnumControlType>();
        Dictionary<string, EnumQuestionType> _QuestionTypeDictE = new Dictionary<string, EnumQuestionType>();


        public ABCardQuestionsForm()
        {
            InitializeComponent();
            _GroupList = new List<string>();
            _ABCardQuestionDataManager= new ABCardQuestionDataManager ();
            //_ControlTypeDictE.Add("", EnumControlType.CHECKBOX);
            //_ControlTypeDictE.Add("", EnumControlType.COMBOBOX);
            //_ControlTypeDictE.Add("", EnumControlType.GRID_COMBOBOX);
            //_ControlTypeDictE.Add("", EnumControlType.GRID_TEXTBOX);
            //_ControlTypeDictE.Add("", EnumControlType.GRID_TEXTBOXDROPDOWN);
            //_ControlTypeDictE.Add("", EnumControlType.RADIO_BUTTON);
            //_ControlTypeDictE.Add("", EnumControlType.TEXTBOX);
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ABCardQuestionsForm_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
            // 載入預先資料
            PrepareData();
        }

        private void cbxGroup_SelectedIndexChanged(object sender, EventArgs e)
        {
            itmPnlNameList.SuspendLayout();
            itmPnlNameList.Items.Clear();

            if (_GroupList.Contains(cbxGroup.Text))
            {
                Dictionary<string, QuestionData> dataDict = _ABCardQuestionDataManager.GetQuestionDataByGroupName(cbxGroup.Text);

                foreach (KeyValuePair<string, QuestionData> data in dataDict)
                {
                    ButtonItem item = new ButtonItem();
                    item.Name = data.Key;
                    item.Text = data.Key;
                    item.Tag = data.Value;
                    item.ButtonStyle = eButtonStyle.ImageAndText;
                    item.OptionGroup = "itemPnlQuestion";
                    item.Click += new EventHandler(item_Click);
                    itmPnlNameList.Items.Add(item);
                }
            }
            SetDefault();
            itmPnlNameList.ResumeLayout();
            itmPnlNameList.Refresh();
        }

        public void controlEnable()
        {
            cbxControlType.Enabled = true;
            cbxQuestionType.Enabled = true;
            chkCanPrint.Enabled = true;
            chkCanStudentEdit.Enabled = true;
            chkCanTeacherEdit.Enabled = true;
            dgItemEnable(true);
            dgItems.Rows.Clear();
            intDisplayOrder.Enabled = true;
        }

        public void dgItemEnable(bool tf)
        {
            if (tf)
            {
                dgItems.Enabled = true;
                dgItems.AllowUserToAddRows = true;
                dgItems.BackgroundColor = Color.White;
            }
            else
            {
                dgItems.Enabled = false;
                dgItems.AllowUserToAddRows = false;
                dgItems.BackgroundColor = Color.Gray;
            }
        }

        void item_Click(object sender, EventArgs e)
        {
            if (itmPnlNameList.SelectedItems.Count > 0)
            {
                controlEnable();
                ButtonItem item = itmPnlNameList.SelectedItems[0] as ButtonItem;
                if (item != null)
                {
                    QuestionData data = item.Tag as QuestionData;
                    _QuestionData = data;
                    if(data !=null)
                    {
                        gp1.Text = data.Name;
                        chkCanPrint.Checked = data.CanPrint;
                        chkCanTeacherEdit.Checked = data.CanTeacherEdit;
                        chkCanStudentEdit.Checked = data.CanStudentEdit;
                        cbxControlType.Text = data.ControlType;
                        cbxQuestionType.Text = data.QuestionType;

                        if (data.QControlType == EnumControlType.GRID_TEXTBOX || data.QControlType == EnumControlType.TEXTBOX)
                            dgItemEnable(false);
                        else
                            dgItemEnable(true);

                        //本人概況_原住民血統項目不提供編輯
                        if (data.Group == "本人概況" && data.Name == "原住民血統")
                            dgItemEnable(false);

                   
                        // 當 checkbox radio 才可以使用可備註
                        if (data.QControlType == EnumControlType.CHECKBOX || data.QControlType == EnumControlType.RADIO_BUTTON)
                            dgItems.Columns[colHasRemark.Index].Visible = true;
                        else
                            dgItems.Columns[colHasRemark.Index].Visible = false;

                        if (data.DisplayOrder.HasValue)
                            intDisplayOrder.Value = data.DisplayOrder.Value;
                        else
                            intDisplayOrder.IsEmpty = true;

                        if (data.itemList.Count > 0)
                        {
                            int rowidx;
                            foreach (QuestionItem qi in data.itemList)
                            {
                                rowidx = dgItems.Rows.Add();
                                dgItems.Rows[rowidx].Cells[colName.Index].Value = qi.Key;
                                dgItems.Rows[rowidx].Cells[colHasRemark.Index].Value = qi.hasRemark;
                            }                        
                        }
                        // 存、歿特殊處理
                        if (data.Name == "直系血親_存、歿")
                        {
                            dgItems.Rows.Clear();
                            dgItemEnable(false);
                        }
                    }
                }               

            }
        }

        

        /// <summary>
        /// 與先載入資料
        /// </summary>
        private void PrepareData()
        {
            _GroupList.Add("本人概況");
            _GroupList.Add("家庭狀況");
            _GroupList.Add("學習狀況");
            _GroupList.Add("自傳");
            _GroupList.Add("自我認識");
            _GroupList.Add("生活感想");
            _GroupList.Add("畢業後計畫");
            _GroupList.Add("備註");
            _GroupList.Add("適應情形");

            // 放入群組控制項
            cbxGroup.Items.AddRange(_GroupList.ToArray());

            // 清除 itemPan
            itmPnlNameList.Items.Clear();
            cbxControlType.Items.Add(EnumControlType.CHECKBOX.ToString());
            cbxControlType.Items.Add(EnumControlType.COMBOBOX.ToString());
            cbxControlType.Items.Add(EnumControlType.GRID_COMBOBOX.ToString());
            cbxControlType.Items.Add(EnumControlType.GRID_TEXTBOX.ToString());
            cbxControlType.Items.Add(EnumControlType.GRID_TEXTBOXDROPDOWN.ToString());
            cbxControlType.Items.Add(EnumControlType.RADIO_BUTTON.ToString());
            cbxControlType.Items.Add(EnumControlType.TEXTBOX.ToString());

            cbxQuestionType.Items.Add(EnumQuestionType.MULTI_ANSWER.ToString());
            cbxQuestionType.Items.Add(EnumQuestionType.PRIORITY.ToString());
            cbxQuestionType.Items.Add(EnumQuestionType.RELATIVE.ToString());
            cbxQuestionType.Items.Add(EnumQuestionType.SEMESTER.ToString());
            cbxQuestionType.Items.Add(EnumQuestionType.MULTI_ANSWER.ToString());
            cbxQuestionType.Items.Add(EnumQuestionType.SIBLING.ToString());
            cbxQuestionType.Items.Add(EnumQuestionType.SINGLE_ANSWER.ToString());
            cbxQuestionType.Items.Add(EnumQuestionType.YEARLY.ToString());
            SetDefault();

        }

        private void SetDefault()
        {
            cbxControlType.Text = "";            
            cbxQuestionType.Text = "";
            intDisplayOrder.IsEmpty = true;
            chkCanPrint.Checked = false;
            chkCanStudentEdit.Checked = false;
            chkCanTeacherEdit.Checked = false;
            dgItems.Rows.Clear();

           
            cbxControlType.Enabled = false;
            cbxQuestionType.Enabled = false;
            intDisplayOrder.Enabled = false;
            chkCanPrint.Enabled = false;
            chkCanStudentEdit.Enabled = false;
            chkCanTeacherEdit.Enabled = false;
            dgItemEnable(false);
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (_QuestionData!=null)
            {
                _QuestionData.ControlType = cbxControlType.Text;
                _QuestionData.QuestionType = cbxQuestionType.Text;
                _QuestionData.DisplayOrder = intDisplayOrder.Value;
                _QuestionData.CanPrint = chkCanPrint.Checked;
                _QuestionData.CanStudentEdit = chkCanStudentEdit.Checked;
                _QuestionData.CanTeacherEdit = chkCanTeacherEdit.Checked;              

                if (dgItems.Rows.Count > 0)
                {
                    List<QuestionItem> qiList = new List<QuestionItem>();
                    foreach (DataGridViewRow drv in dgItems.Rows)
                    {
                        if (drv.IsNewRow)
                            continue;
                        bool hasR = false;

                        if(drv.Cells[colHasRemark.Index].Value !=null)
                            hasR=bool.Parse(drv.Cells[colHasRemark.Index].Value.ToString());

                        QuestionItem qi;
                        string key = drv.Cells[colName.Index].Value.ToString();
                        if (hasR)
                            qi = new QuestionItem(key, hasR);
                        else
                            qi = new QuestionItem(key);

                        qiList.Add(qi);
                    }
                    _QuestionData.itemList = qiList;
                }

                List<UDTQuestionsDataDef> data = new List<UDTQuestionsDataDef>();
                data.Add(_QuestionData.GetUpdateData());
                UDTTransfer.ABUDTQuestionsDataUpdate(data);
                // 呼叫同步
                EventHub.OnCounselChanged();
                FISCA.Presentation.Controls.MsgBox.Show("儲存完成!");                
            }
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            if (FISCA.Presentation.Controls.MsgBox.Show("將清除所有題目與選項恢復到系統預設值，執行後請重新登入系統，請問確認執行?", "清空題目", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.Yes)
            {
                List<UDTQuestionsDataDef> list = UDTTransfer.ABUDTQuestionsDataSelectAll();
                UDTTransfer.ABUDTQuestionsDataDelete(list);
            }
        }
    }
}
