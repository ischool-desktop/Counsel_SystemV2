using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Counsel_System2.UI
{
    public partial class QuestionGroupUIDefinition : UserControl
    {
        private Parser.QuestionGroup qgroup;
        public QuestionGroupUIDefinition()
        {
            InitializeComponent();
        }

        public void SetQuestionGroup(Parser.QuestionGroup qgroup)
        {
            this.qgroup = qgroup;
            this.textBoxX1.Text = qgroup.GetLabel();

            fillGrid(qgroup);
        }

        private void fillGrid(Parser.QuestionGroup qgroup)
        {
            this.dataGridViewX1.Rows.Clear();

            List<Parser.QuestionListItem> items = qgroup.GetListItems();
            foreach (Parser.QuestionListItem item in items)
            {
                object[] obj = { item.GetLabel(), item.Selected, item.HasText };
                this.dataGridViewX1.Rows.Add(obj);
            }
        }

        private void labelX2_Click(object sender, EventArgs e)
        {

        }

        private void labelX2_MouseHover(object sender, EventArgs e)
        {
            
        }

        private void QuestionGroupUIDefinition_Load(object sender, EventArgs e)
        {
            //this.balloonTip1.SetBalloonCaption(this.labelX2, "");
            this.balloonTip1.SetBalloonText(this.labelX2, "以下的清單選項可供此題組下所有題目共用！");
            this.balloonTip1.ShowBalloon(this.labelX2);
        }

        public void UpdateNewValue()
        {
            this.qgroup.SetLabel(this.textBoxX1.Text);
            List<Parser.QuestionListItem> items = this.qgroup.GetListItems();
            items.Clear();

            foreach (DataGridViewRow row in this.dataGridViewX1.Rows)
            {
                if (!row.IsNewRow)
                {
                    Parser.QuestionListItem item = new Parser.QuestionListItem();
                    item.SetLabel((row.Cells[0].Value == null) ? "" : row.Cells[0].Value.ToString());
                    item.Selected = (row.Cells[1].Value==null) ? false : (bool)row.Cells[1].Value;
                    item.HasText = (row.Cells[2].Value ==null) ? false : (bool)row.Cells[2].Value;
                    items.Add(item);
                }
            }
        }

        private void dataGridViewX1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
        } // end function

        private void SetCellToModify(string msg, DataGridViewCellValidatingEventArgs e)
        {
            MessageBox.Show(msg);
            e.Cancel = true;
        }

        private void dataGridViewX1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            //去掉 new row
            DataGridViewRow rowX = this.dataGridViewX1.Rows[e.RowIndex];
            if (rowX.IsNewRow)
                return;

            //是第一欄才判斷
            if (e.ColumnIndex == 0)
            {                
                if ((e.FormattedValue == null) ||
                        (string.IsNullOrEmpty(e.FormattedValue.ToString())))
                {
                    SetCellToModify("請輸入選項值", e);
                    //e.Cancel = true;
                }
                else
                {
                    //檢查是否有重複選項值
                    for (int i = 0; i < this.dataGridViewX1.Rows.Count; i++)
                    {
                        if (i != e.RowIndex)
                        {
                            DataGridViewRow row = this.dataGridViewX1.Rows[i];

                            if (!row.IsNewRow &&
                                    (row.Cells[0].Value.ToString() == e.FormattedValue.ToString()))
                            {
                                SetCellToModify(string.Format("您輸入的值：{0} 和 第 {1} 列的值重複！", e.FormattedValue.ToString(), (i + 1).ToString()), e);
                                //e.Cancel = true;
                                break;
                            }
                        }

                    } //end for

                } // end if
            }// end if
        }

    }// end of class
} // end of package
