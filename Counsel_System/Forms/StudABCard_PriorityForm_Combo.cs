using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Counsel_System2.DAO;

namespace Counsel_System2.Forms
{
    public partial class StudABCard_PriorityForm_Combo : FISCA.Presentation.Controls.BaseForm
    {
        UDTPriorityDataDef _data;

        public StudABCard_PriorityForm_Combo(UDTPriorityDataDef Data)
        {
            InitializeComponent();
            _data = Data;        
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }

        /// <summary>
        /// 取得資料
        /// </summary>
        /// <returns></returns>
        public UDTPriorityDataDef GetData()
        {
            return _data;
        }

        /// <summary>
        /// 設定標頭名稱
        /// </summary>
        /// <param name="name"></param>
        public void SetTitleName(string name)
        {
            this.Text = name;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            SetData();
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }

        private void StudABCard_PriorityForm_Combo_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
            LoadDataToUI();
        }

        private void LoadDataToUI()
        {
            cbx01.Text = _data.P1;
            cbx02.Text = _data.P2;
            cbx03.Text = _data.P3;
        }

        private void SetData()
        {
            _data.P1 = cbx01.Text;
            _data.P2 = cbx02.Text;
            _data.P3 = cbx03.Text;
        }

        public void SetItemList(List<string> items)
        {
            cbx01.Items.Clear();
            cbx02.Items.Clear();
            cbx03.Items.Clear();
            cbx01.Items.AddRange(items.ToArray());
            cbx02.Items.AddRange(items.ToArray());
            cbx03.Items.AddRange(items.ToArray());
        }
    }
}
