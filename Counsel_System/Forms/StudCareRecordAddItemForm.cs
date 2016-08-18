using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Counsel_System2.Forms
{
    public partial class StudCareRecordAddItemForm : FISCA.Presentation.Controls.BaseForm
    {
        List<string> _itemList;
        DAO.CareRecordItemManager.ItemType _itemType;
        private string _itemName = "";
        public StudCareRecordAddItemForm(List<string> itemList,DAO.CareRecordItemManager.ItemType itmType)
        {
            InitializeComponent();
            this.MaximumSize = this.MinimumSize = this.Size;
            _itemList = itemList;
            _itemType = itmType;
        }

        private void StudCareRecordAddItemForm_Load(object sender, EventArgs e)
        {

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {            
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            // 檢查是否已經有資料

            _itemName = txtName.Text;
            if (string.IsNullOrWhiteSpace(txtName.Text))
            {
                FISCA.Presentation.Controls.MsgBox.Show("請輸入名稱.");
                return;
            }

            if (_itemList.Contains(_itemName))
            {
                FISCA.Presentation.Controls.MsgBox.Show("名稱已存在無法新增.");
                return;
            }
            // 新增項目
            Global._CareRecordItemManager.AddItem(_itemName, _itemType);
            
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }

        /// <summary>
        /// 取得新增項目名稱
        /// </summary>
        /// <returns></returns>
        public string GetAddItemName()
        {
            return _itemName;
        }
    }
}
