using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Counsel_System2.DAO;
using DevComponents.DotNetBar.Controls;

namespace Counsel_System2.Forms
{
    public partial class StudABCard_YearlyForm_TextDropDown : Form
    {
        UDTYearlyDataDef _data;
        List<CheckedListBox> _CheckList = new List<CheckedListBox>();
        Dictionary<string, TextBoxDropDown> _TextBoxDropDownDic;
               
        public StudABCard_YearlyForm_TextDropDown(UDTYearlyDataDef data)
        {
            InitializeComponent();
            _data = data;
            _TextBoxDropDownDic = new Dictionary<string, TextBoxDropDown>();
            foreach (Control cr in this.panelEx1.Controls)
                if (cr is TextBoxDropDown)
                {
                    TextBoxDropDown dd = cr as TextBoxDropDown;
                    _TextBoxDropDownDic.Add(cr.Name,dd);
                }
             
            foreach(string key in _TextBoxDropDownDic.Keys)
             {
                CheckedListBox itm = new CheckedListBox();
                itm.Name = key;
                itm.LostFocus += new EventHandler(itm_LostFocus);
                _CheckList.Add(itm);
            }            
        }

        void itm_LostFocus(object sender, EventArgs e)
        {
            CheckedListBox itm = sender as CheckedListBox;
            if(_TextBoxDropDownDic.ContainsKey(itm.Name))
            {
                List<string> tmp = new List<string> ();
                foreach(object  cb in itm.CheckedItems)
                    tmp.Add(cb.ToString());
                _TextBoxDropDownDic[itm.Name].Text=string.Join(",",tmp.ToArray());

            }
        }

        private void StudABCard_YearlyForm_TextDropDown_Load(object sender, EventArgs e)
        {
            this.MinimumSize = this.MaximumSize = this.Size;
            LoadData();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            SetData();
            this.Close();
        }

        public UDTYearlyDataDef GetData()
        {
            SetData();
            return _data;
        }

        public void SetTitleName(string name)
        {
            this.Text = name;
        }

        private void LoadData()
        {
            txtDP01.Text = _data.G1;
            txtDP02.Text = _data.G2;
            txtDP03.Text = _data.G3;
            txtDP04.Text = _data.G4;
            txtDP05.Text = _data.G5;
            txtDP06.Text = _data.G6;

            List<string> tmpStr = new List<string>();
            // 解析並且放入
            foreach(CheckedListBox chk in _CheckList)
            {
                if (_TextBoxDropDownDic.ContainsKey(chk.Name))
                {
                    tmpStr.Clear();
                    tmpStr = _TextBoxDropDownDic[chk.Name].Text.Split(',').ToList();
                    foreach(string key in tmpStr)
                    foreach (CheckBoxX ck in chk.Items)
                    {
                        if (ck.Text == key)
                            ck.Checked = true;
                    }
                }

            }
        }

        private void SetData()
        {
            _data.G1 = txtDP01.Text;
            _data.G2 = txtDP02.Text;
            _data.G3 = txtDP03.Text;
            _data.G4 = txtDP04.Text;
            _data.G5 = txtDP05.Text;
            _data.G6 = txtDP06.Text;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;

            this.Close();
        }

        public void SetSelectItems(List<string> items)
        {
            foreach (CheckedListBox chk in _CheckList)
            {
                chk.Items.Clear();
                chk.Items.AddRange(items.ToArray());
                if (_TextBoxDropDownDic.ContainsKey(chk.Name))
                    _TextBoxDropDownDic[chk.Name].DropDownControl = chk;
            }
        }


    }
}
