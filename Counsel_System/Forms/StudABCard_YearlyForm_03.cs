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
    public partial class StudABCard_YearlyForm_03 : FISCA.Presentation.Controls.BaseForm
    {
        Dictionary<string, UDTYearlyDataDef> _data;
        int _GradeYear = 0;
        Dictionary<string, int> _gYear = new Dictionary<string, int>();
        // 計算位置使用
        int y = 0;
        int labHeight = 23;

        string key1 = "學習狀況_最喜歡的學科";
        string key2 = "學習狀況_最感困難的學科";
        string key3 = "學習狀況_特殊專長";
        string key4 = "學習狀況_休閒興趣";
        string key5 = "學習狀況_社團幹部";
        string key6 = "學習狀況_班級幹部";
        string key7 = "學習狀況_特殊專長_樂器演奏";
        string key8 = "學習狀況_特殊專長_外語能力";


        public StudABCard_YearlyForm_03(Dictionary<string,UDTYearlyDataDef> data,string gradeyearStr)
        {
            InitializeComponent();
            _gYear.Add("一年級", 1);
            _gYear.Add("二年級", 2);
            _gYear.Add("三年級", 3);
            _gYear.Add("四年級", 4);
            _gYear.Add("五年級", 5);
            _gYear.Add("六年級", 6);            
            _data = data;

            this.Text = gradeyearStr + "學習狀況";
            if (_gYear.ContainsKey(gradeyearStr))
                _GradeYear = _gYear[gradeyearStr];            
        }

        public void SetFlp01Items(List<string> nameList)
        {
            addFlpItems(flp01, nameList);
        }

        private void addFlpItems(FlowLayoutPanel flp, List<string> nameList)
        {
            foreach (string name in nameList)
            {
                CheckBox cb = new CheckBox();
                cb.Name = name;
                cb.Text = name;
                cb.Checked = false;
                cb.AutoSize = true;
                flp.Controls.Add(cb);
            }
        }

        public void SetFlp02Items(List<string> nameList)
        {
            addFlpItems(flp02, nameList);
        }

        public void LoadData()
        {
            // 最喜歡的學科
            if (_data.ContainsKey(key1))
                txt01.Text = GetItemString(_data[key1]);

            // 最感困難的學科
            if (_data.ContainsKey(key2))
                txt02.Text = GetItemString(_data[key2]);

            // 特殊專長樂器演奏
            if (_data.ContainsKey(key7))
                textBoxX2.Text = GetItemString(_data[key7]);

            // 特殊專長外語能力
            if (_data.ContainsKey(key8))
                textBoxX1.Text = GetItemString(_data[key8]);


            // 特殊專長
            if (_data.ContainsKey(key3))
            {
                List<string> items = new List<string> ();
                items =GetItemString(_data[key3]).Split(',').ToList();

                foreach (Control  ctr in flp01.Controls)
                {
                    if (ctr is CheckBox)
                        if (items.Contains(ctr.Name))
                        {
                            CheckBox cb = ctr as CheckBox;
                            cb.Checked = true;
                        }
                }
            }

            // 休閒興趣
            if (_data.ContainsKey(key4))
            {
                List<string> items = new List<string>();
                items = GetItemString(_data[key4]).Split(',').ToList();

                foreach (Control ctr in flp02.Controls)
                {
                    if (ctr is CheckBox)
                        if (items.Contains(ctr.Name))
                        {
                            CheckBox cb = ctr as CheckBox;
                            cb.Checked = true;
                        }
                }
            }

            //// 社團幹部
            //if(_data.ContainsKey(key5))
            //    txt03.Text = GetItemString(_data[key5]);

            //// 班級幹部
            //if (_data.ContainsKey(key6))
            //    txt04.Text = GetItemString(_data[key6]);

        }

        // 取得資料
        private string GetItemString(UDTYearlyDataDef data)
        {
            string retVal = "";
            
            switch (_GradeYear)
            {
                case 1: retVal = data.G1; break;
                case 2: retVal = data.G2; break;
                case 3: retVal = data.G3; break;
                case 4: retVal = data.G4; break;
                case 5: retVal = data.G5; break;
                case 6: retVal = data.G6; break;
            }
            return retVal;
        }

        // 設定資料
        private void SetItemString(string key, string value)
        {
            if (_data.ContainsKey(key))
            {
                switch (_GradeYear)
                {
                    case 1: _data[key].G1 = value; break;
                    case 2: _data[key].G2 = value; break;
                    case 3: _data[key].G3 = value; break;
                    case 4: _data[key].G4 = value; break;
                    case 5: _data[key].G5 = value; break;
                    case 6: _data[key].G6 = value; break;
                }
            }
        }
        
        // 回存資料
        private void SetData()
        {
            // 最喜歡的學科
            SetItemString(key1, txt01.Text);

            // 最感困難的學科
            SetItemString(key2, txt02.Text);

            //樂器演奏
            SetItemString(key7, textBoxX2.Text);

            // 外語能力
            SetItemString(key8, textBoxX1.Text);

            // 特殊專長
            List<string> item1 = new List<string>();
            foreach (Control cr in flp01.Controls)
            {
                if (cr is CheckBox)
                {
                    CheckBox cb = cr as CheckBox;
                    if (cb.Checked)
                        item1.Add(cb.Name);
                }
            }
            SetItemString(key3, string.Join(",", item1.ToArray()));
                
            // 休閒興趣
            List<string> item2 = new List<string>();
            foreach (Control cr in flp02.Controls)
            {
                if (cr is CheckBox)
                {
                    CheckBox cb = cr as CheckBox;
                    if (cb.Checked)
                        item2.Add(cb.Name);
                }
            }
            SetItemString(key4, string.Join(",", item2.ToArray()));

            //// 社團幹部
            //SetItemString(key5, txt03.Text);

            //// 班級幹部
            //SetItemString(key6, txt04.Text);
        }

        public Dictionary<string, UDTYearlyDataDef> GetData()
        {
            return _data;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            SetData();
            this.Close();
        }

        private void StudABCard_YearlyForm_03_Load(object sender, EventArgs e)
        {
            try
            {
                LoadData();
            }
            catch (Exception ex)
            { 
            
            }
            setUI();
        }

        private void setUI()
        {
            int xPos = lbl01.Location.X;

            y += lbl01.Location.Y+labHeight;
            flp01.Location = new Point(flp01.Location.X, y);
            flp01.BorderStyle = BorderStyle.FixedSingle;
            y += flp01.Size.Height;

            y += 12;
            labelX2.Location = new Point(xPos, y);
            y += labelX2.Height;
            textBoxX2.Location = new Point(xPos, y);
            y += textBoxX2.Height;

            y += 12;
            labelX1.Location = new Point(xPos, y);
            y += labelX1.Height;
            textBoxX1.Location = new Point(xPos, y);
            y += textBoxX1.Height;

            y += 12;
            lbl02.Location = new Point(lbl02.Location.X, y);
            y += labHeight;
            flp02.Location = new Point(flp02.Location.X, y);
            flp02.BorderStyle = BorderStyle.FixedSingle;
            y += flp02.Size.Height;
            
            y += 12;

            lbl03.Location = new Point(lbl03.Location.X, y);
            y += labHeight;
            txt01.Location = new Point(txt01.Location.X, y);
            y += txt01.Size.Height;
            y += 12;

            lbl04.Location = new Point(lbl04.Location.X, y);
            y += labHeight;
            txt02.Location = new Point(txt02.Location.X, y);
            y += txt02.Size.Height;
            y += 12;


            //lbl05.Location = new Point(lbl05.Location.X, y);
            //y += labHeight;
            //txt03.Location = new Point(txt03.Location.X, y);
            //y += txt03.Size.Height;
            //y += 12;

            //lbl06.Location = new Point(lbl06.Location.X, y);
            //y += labHeight;
            //txt04.Location = new Point(txt04.Location.X, y);
            //y += txt04.Size.Height;
            y += 12;

            

            this.Size = new Size(this.Width, y+70 );
            this.MinimumSize = this.MaximumSize = this.Size;
        }
    }
}
