using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using K12.Data;

namespace Counsel_System2.Forms
{
    public partial class StudCareRecordForm : FISCA.Presentation.Controls.BaseForm
    {
        DAO.UDT_CounselCareRecordDef _CounselCareRecord;
        public enum accessType{Insert,Edit}
        accessType _accessType;
        DAO.UDTTransfer _UDTTransfer;
        DAO.LogTransfer _LogTransfer;
        StudentRecord _studRec;
        string insertKey = "新增..";
        List<string> _itemListCategory;
        List<string> _itemListOrigin;

        public StudCareRecordForm(DAO.UDT_CounselCareRecordDef CareRecordDef,accessType accType)
        {
            InitializeComponent();
            this.MaximumSize = this.MinimumSize = this.Size;
            _accessType = accType;
            _CounselCareRecord = CareRecordDef;
            _itemListCategory = new List<string>();
            _itemListOrigin = new List<string>();

            _studRec = Student.SelectByID(_CounselCareRecord.StudentID.ToString());
            _UDTTransfer = new DAO.UDTTransfer();
            _LogTransfer = new DAO.LogTransfer();
            //if (accType == accessType.Insert)
            //    _CounselCareRecord.AuthorID = Utility.GetAuthorID();
        }

        private void LoadDefaultDataToForm()
        {
            LoadItemData();
        
        }

        private void LoadItemData()
        {
            _itemListCategory.Clear();            
            _itemListCategory = Global._CareRecordItemManager.GetItemList(DAO.CareRecordItemManager.ItemType.個案類別);
            _itemListCategory.Add(insertKey);
            cbxCaseCategory.Items.Clear();
            cbxCaseCategory.Items.AddRange(_itemListCategory.ToArray());

            _itemListOrigin.Clear();
            _itemListOrigin = Global._CareRecordItemManager.GetItemList(DAO.CareRecordItemManager.ItemType.個案來源);
            _itemListOrigin.Add(insertKey);
            cbxCaseOrigin.Items.Clear();
            cbxCaseOrigin.Items.AddRange(_itemListOrigin.ToArray());

        }

        private void LoadUDTDataToForm()
        {
            
            if (_studRec != null)
            {
                if (_studRec.Class != null)
                {
                    if (_studRec.Class.GradeYear.HasValue)
                        lblGradeYear.Text = _studRec.Class.GradeYear.Value.ToString();

                    lblClassName.Text = _studRec.Class.Name;
                    lblName.Text = _studRec.Name;
                }
            }
            txtCodeName.Text = _CounselCareRecord.CodeName;
            if(_CounselCareRecord.FileDate.HasValue )
                dtFileDate.Value = _CounselCareRecord.FileDate.Value;

            cbxCaseCategory.Text = _CounselCareRecord.CaseCategory;

            txtCaseCategoryRemark.Text = _CounselCareRecord.CaseCategoryRemark;

            cbxCaseOrigin.Text = _CounselCareRecord.CaseOrigin;

            txtCaseOriginRemark.Text = _CounselCareRecord.CaseOriginRemark;
            txtSuperiority.Text = _CounselCareRecord.Superiority;
            txtWeakness.Text = _CounselCareRecord.Weakness;
            txtOtherinstitute.Text = _CounselCareRecord.OtherInstitute;
            txtCounselGoal.Text = _CounselCareRecord.CounselGoal;
            txtCounselType.Text = _CounselCareRecord.CounselType;
            txtAssistedMatter.Text = _CounselCareRecord.AssistedMatter;
            txtAuthor_id.Text = _CounselCareRecord.AuthorID;
            txtAuthorName.Text = _CounselCareRecord.AuthorName;

            _LogTransfer.Clear();
            LogData();
            if(string.IsNullOrEmpty(txtAuthor_id.Text))
                txtAuthor_id.Text = Utility.GetAuthorID();
        }

        private bool CheckData()
        {
            bool retVal = true;
            if (string.IsNullOrEmpty(txtAuthor_id.Text) || string.IsNullOrEmpty(cbxCaseCategory.Text) || string.IsNullOrEmpty(cbxCaseOrigin.Text) || string.IsNullOrEmpty(dtFileDate.Text))
                retVal = false;                

            return retVal;
        }

        private void SaveFormDataToUDTDef()
        {
            _CounselCareRecord.CodeName = txtCodeName.Text;
            _CounselCareRecord.FileDate = dtFileDate.Value;
            _CounselCareRecord.CaseCategory = cbxCaseCategory.Text;

            _CounselCareRecord.CaseCategoryRemark = txtCaseCategoryRemark.Text;

            _CounselCareRecord.CaseOrigin = cbxCaseOrigin.Text;
            _CounselCareRecord.CaseOriginRemark = txtCaseOriginRemark.Text;
            _CounselCareRecord.Superiority = txtSuperiority.Text;
            _CounselCareRecord.Weakness = txtWeakness.Text;
            _CounselCareRecord.OtherInstitute = txtOtherinstitute.Text;
            _CounselCareRecord.CounselGoal = txtCounselGoal.Text;
            _CounselCareRecord.CounselType = txtCounselType.Text;
            _CounselCareRecord.AssistedMatter = txtAssistedMatter.Text;
            _CounselCareRecord.AuthorID = txtAuthor_id.Text;
            _CounselCareRecord.AuthorName = txtAuthorName.Text;
           
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void StudCareRecordForm_Load(object sender, EventArgs e)
        {
            LoadDefaultDataToForm();
            LoadUDTDataToForm();          
            
            txtCaseCategoryRemark.Enabled = false;
            txtCaseOriginRemark.Enabled = false;
            ChecktxtRemarkEnable();
        }

        private void LogData()
        {
            _LogTransfer.SetLogValue("代號", _CounselCareRecord.CodeName);
            if (_CounselCareRecord.FileDate.HasValue)
                _LogTransfer.SetLogValue("立案日期", _CounselCareRecord.FileDate.Value.ToShortDateString());
            else
                _LogTransfer.SetLogValue("立案日期", "");

            _LogTransfer.SetLogValue("個案類別", _CounselCareRecord.CaseCategory);
            _LogTransfer.SetLogValue("個案類別備註", _CounselCareRecord.CaseCategoryRemark);
            _LogTransfer.SetLogValue("個案來源", _CounselCareRecord.CaseOrigin);
            _LogTransfer.SetLogValue("個案來源備註", _CounselCareRecord.CaseOriginRemark);
            _LogTransfer.SetLogValue("優勢能力及財力", _CounselCareRecord.Superiority);
            _LogTransfer.SetLogValue("弱勢能力及阻力", _CounselCareRecord.Weakness);
            _LogTransfer.SetLogValue("輔導人員輔導目標", _CounselCareRecord.CounselGoal);
            _LogTransfer.SetLogValue("校外協輔機構", _CounselCareRecord.OtherInstitute);
            _LogTransfer.SetLogValue("輔導人員輔導方式", _CounselCareRecord.CounselType);
            _LogTransfer.SetLogValue("協同輔導人員協助導師事項", _CounselCareRecord.AssistedMatter);

            _LogTransfer.SetLogValue("記錄者", _CounselCareRecord.AuthorID);
            _LogTransfer.SetLogValue("記錄者姓名", _CounselCareRecord.AuthorName);
            _LogTransfer.SetLogValue("個案類別", _CounselCareRecord.CaseCategory);
        }


        private void btnSave_Click(object sender, EventArgs e)
        {
            if (CheckData() == false)
            {
                Utility.ShowCannotSaveMessage();
                return;
            }
            SaveFormDataToUDTDef();

            // log
            LogData();

            string studStr = "學號:" + _studRec.StudentNumber + ",姓名:" + _studRec.Name + ",";
            if (_accessType == accessType.Insert)
            {
                // 檢查是否可以新增 (立案日期+個案類別 不能重複)
                List<DAO.UDT_CounselCareRecordDef> CareRecList = _UDTTransfer.GetCareRecordsByStudentID(_studRec.ID);               
                bool pass = true;
                foreach (DAO.UDT_CounselCareRecordDef rec in CareRecList)
                {
                    if (rec.FileDate.HasValue && _CounselCareRecord.FileDate.HasValue)                    
                        if (rec.FileDate.Value.ToShortDateString() == _CounselCareRecord.FileDate.Value.ToShortDateString())                        
                            if (rec.CaseCategory == _CounselCareRecord.CaseCategory)
                                pass = false;                    
                }

                if (pass)
                {
                    // log
                    _LogTransfer.SaveInsertLog("學生.優先關懷紀錄-新增", "新增", studStr, "", "student", _studRec.ID);
                    _UDTTransfer.InsertCareRecord(_CounselCareRecord);
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("已有相同的立案日期與個案類別，無法新增。");
                    return;
                }
            }
            else
            {
                // log
                _LogTransfer.SaveChangeLog("學生.優先關懷紀錄-修改", "修改", studStr, "", "student", _studRec.ID);

                _UDTTransfer.UpdateCareRecord(_CounselCareRecord);
            }

            Utility.ShowSavedMessage();
            this.DialogResult = System.Windows.Forms.DialogResult.OK;            
        }

        private void cbxCaseCategory_SelectedIndexChanged(object sender, EventArgs e)
        {
            ChecktxtRemarkEnable();
            if (cbxCaseCategory.Text == insertKey)
            {
                StudCareRecordAddItemForm scraf = new StudCareRecordAddItemForm(_itemListCategory, DAO.CareRecordItemManager.ItemType.個案類別);
                if (scraf.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string itemName = scraf.GetAddItemName();
                    LoadItemData();
                    cbxCaseCategory.Text = itemName;
                }
                cbxCaseCategory.Text = null;
            }
        }

        /// <summary>
        /// 檢查當選擇是其它，備註功能開啟，否則關閉。
        /// </summary>
        private void ChecktxtRemarkEnable()
        {
            if (cbxCaseCategory.Text=="其它")
                txtCaseCategoryRemark.Enabled = true;
            else
            {
                txtCaseCategoryRemark.Enabled = false;
                txtCaseCategoryRemark.Text = "";
            }
            if (cbxCaseOrigin.Text=="其它")
                txtCaseOriginRemark.Enabled = true;
            else
            {
                txtCaseOriginRemark.Enabled = false;
                txtCaseOriginRemark.Text = "";
            }
        }

        private void cbxCaseOrigin_SelectedIndexChanged(object sender, EventArgs e)
        {
            ChecktxtRemarkEnable();
            if (cbxCaseOrigin.Text == insertKey)
            {
                StudCareRecordAddItemForm scraf = new StudCareRecordAddItemForm(_itemListOrigin, DAO.CareRecordItemManager.ItemType.個案來源);
                if (scraf.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string itemName = scraf.GetAddItemName();
                    LoadItemData();
                    cbxCaseOrigin.Text = itemName;
                }
                else
                    cbxCaseOrigin.Text = null;
                
            }
        }
    }
}
