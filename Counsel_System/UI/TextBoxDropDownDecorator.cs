using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DevComponents.DotNetBar.Controls;
using System.Windows.Forms;

namespace Counsel_System2.UI
{
    class TextBoxDropDownDecorator
    {
         private TextBoxDropDown txt;
        private CheckedListBox chkBox;
        private List<string> otherValues;   //儲存除了 checkedlistbox.Items 以外的輸入值，以免 lost

        public TextBoxDropDownDecorator(CheckedListBox chkBox)
        {
            this.chkBox = chkBox;
            this.otherValues = new List<string>();
            init();
        }

        public void SetTextBox(TextBoxDropDown txt)
        {
            this.txt = txt;
            this.txt.DropDownControl = this.chkBox;
            this.otherValues.Clear();

            //Find values do not exist in the CheckedListBox.Items 
            Dictionary<string, int> items = new Dictionary<string, int>();
            for (int i=0; i<this.chkBox.Items.Count ; i++)
            {
                object obj = this.chkBox.Items[i];
                items.Add(obj.ToString().Trim(), i);
                this.chkBox.SetItemChecked(i, false);
            }

            if (!string.IsNullOrEmpty(this.txt.Text.Trim()))
            {
                string[] values = this.txt.Text.Split(new char[] { ',' });
                foreach (string str in values)
                {
                    if (!string.IsNullOrEmpty(str.Trim())) {

                        if (!items.ContainsKey(str.Trim()))
                        {
                            this.otherValues.Add(str.Trim());
                        }
                        else   //有比對到，則 ckeck this item
                        {
                            this.chkBox.SetItemChecked(items[str.Trim()], true);
                        }
                    }
                } // end foreach
            } // end if
        }

        private void init()
        {
            this.chkBox.ItemCheck += new ItemCheckEventHandler(chkBox_ItemCheck);
            this.chkBox.MouseClick += new MouseEventHandler(chkBox_MouseClick);
            this.chkBox.Leave += new EventHandler(chkBox_Leave);
            this.chkBox.LostFocus += new EventHandler(chkBox_LostFocus);
        }

        void chkBox_LostFocus(object sender, EventArgs e)
        {
            this.getValue();
        }

        void chkBox_Leave(object sender, EventArgs e)
        {
            //this.getValue();
        }

        void chkBox_MouseClick(object sender, MouseEventArgs e)
        {
            //this.getValue();
        }

        void getValue()
        {
            string result = "";
            foreach (object obj in this.chkBox.CheckedItems)
            {
                if (this.chkBox.GetItemCheckState(this.chkBox.Items.IndexOf(obj)) == CheckState.Checked)
                    result += obj.ToString() + ", ";
            }

            foreach (string str in this.otherValues)
            {
                result += str + ", ";
            }

            this.txt.Text = result;
        }

        void chkBox_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            //this.getValue();
        }
    }
}
