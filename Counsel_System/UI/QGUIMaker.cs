using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using DevComponents.DotNetBar;
using DevComponents.DotNetBar.Controls;
using System.Xml;
using Counsel_System2.Parser;

namespace Counsel_System2.UI
{
    /// <summary>
    /// This is a decoration class which will Render one QuestionGroup UI，
    /// </summary>
    class QGUIMaker
    {
        // Declare the delegate (if using non-generic pattern).
        public delegate void ContentChange(object sender, EventArgs e);

        // Declare the event.
        public event ContentChange OnContentChange;

        private QGPanel pnlQGroup;
        private QuestionGroup questionGroup;
        private FlowLayoutPanel contentPanel;
        private Answers answers;
        private CheckedListBox chkBox;    //作為 TextBoxDropDown 控制項的選擇清單用。
        private TextBoxDropDownDecorator textBoxDropDownDecorator;  //用來處理 TextBoxDropDown 控制項以及 CheckedListBox 互動的 Decorator 物件

        private string qType = "";
        private Dictionary<string, Control> allQControls;  //紀錄所有題目的控制項，以利後續設定答案時使用。

        public QGUIMaker(QGPanel pnl, QuestionGroup questionGroup)
        {
            this.pnlQGroup = pnl;
            this.questionGroup = questionGroup;
            this.pnlQGroup.FlowDirection = FlowDirection.TopDown;   // why?

            adjHeight();
            
            //Make Header :
            Font font = this.pnlQGroup.Font;
            SizeF stringSize = pnl.CreateGraphics().MeasureString(this.questionGroup.GetLabel(), font);

            TextBox lbl = new TextBox();
            lbl.Multiline = true;
            lbl.BorderStyle = BorderStyle.None;
            lbl.ReadOnly = true;
            lbl.BackColor = Color.White;
            lbl.Width = 80; //default width
            if (this.questionGroup.GetLabelVisible())
            {
                if (!string.IsNullOrEmpty(this.questionGroup.GetLabelWidth()))
                {
                    int width = 0;
                    bool isNumber = int.TryParse(this.questionGroup.GetLabelWidth(), out width);
                    if (isNumber)
                        lbl.Width = width;
                }
            }
            else
                lbl.Width = 0;


            
            //lbl.Width = (this.questionGroup.GetLabelVisible() ? 80 : 0);
            lbl.Height = (lbl.Width == 0) ? 0 : 20 * ((int)stringSize.Width / lbl.Width + 1);
            lbl.Text = this.questionGroup.GetLabel() + "：";
            System.Windows.Forms.Padding margin = lbl.Margin;
            margin.Top = 10;
            lbl.Margin = margin;            
            this.pnlQGroup.Controls.Add(lbl);

            /*
            Label lbl = new Label();
            lbl.Width = (int)stringSize.Width + 20;
            lbl.Height = 25;
            lbl.Top = 4;
            lbl.Text = this.questionGroup.GetLabel() + "：";
            System.Windows.Forms.Padding margin = lbl.Margin;
            margin.Top = 10;
            lbl.Margin = margin;
            //lbl.BackColor = System.Drawing.Color.Red;
            this.pnlQGroup.Controls.Add(lbl);
            */

            //make content panel
            this.contentPanel = new FlowLayoutPanel();
            this.contentPanel.FlowDirection = FlowDirection.LeftToRight;
            this.contentPanel.Height = this.pnlQGroup.Height - lbl.Height + 10;
            this.contentPanel.Width = this.pnlQGroup.Width - lbl.Width - 10;
            this.contentPanel.WrapContents = true;
            //this.contentPanel.BackColor = System.Drawing.Color.Green;
            this.pnlQGroup.Controls.Add(this.contentPanel);

            this.allQControls = new Dictionary<string, Control>();

            makeUI();
        }

        /**
         * <Answers label="本人概況"> <!-- 對應到 Subject的 label -->
	            <Ans name="AAA12345670" value=”AB”></Ans>
	            <Ans name="AAA12345671" value=”基督教”></Ans>	            
	            <Ans name="AAA12345678" >
		            <Item value=”近視” />
                    <Item value=”其它” remark=”蛀牙” />
                </Ans>
	            <Ans name="AAA12345679">
                    <Item value=”氣喘” />
                    <Item value=”其它” remark=”身心症” />
                </Ans>
            </Answers>
         * */
        public void SetAnswer(XmlElement answers)
        {
            this.answers = new Answers(answers);
            foreach (Answer ans in this.answers.GetAllAnswers())
            {
                Question q = this.questionGroup.GetQuestionByName(ans.GetName());   //根據解答名稱找到該題目
                if (q != null)  //answer 是一整包 subject 裡所有題目的答案，而此 questionGroup中只包含其中一部分。
                {

                    if (q.GetQuestionType().ToLower() == "checkbox")   //checkbox 為複選，
                    {
                        foreach (AnswerItem ai in ans.GetAnswerItems())
                        {
                            //找到相關的 checkbox，並設定值為 checked
                            CheckBox chk = null;
                            if (this.allQControls.ContainsKey(ai.GetValueName()))
                            {
                                chk = (CheckBox)this.allQControls[ai.GetValueName()];
                                chk.Checked = true;
                                QuestionListItem questionListItem = q.GetListItemByLable(ai.GetValueName());
                                if (questionListItem != null)
                                {
                                    if (questionListItem.HasText)
                                    {
                                        TextBox txt = (TextBox)this.allQControls[chk.Name + "_remark"];
                                        txt.Text = ai.GetValueRemark();
                                    }
                                }
                            }

                        }
                    }
                    else if (q.GetQuestionType().ToLower() == "grid")
                    {
                        DataGridView dg = (DataGridView)this.allQControls[ans.GetName()];
                        dg.Rows.Clear();

                        foreach (AnswerItem ai in ans.GetAnswerItems())
                        {
                            //每筆 AnswerItem 是一筆 Row
                            List<object> cellValues = new List<object>();
                            Dictionary<string, string> contents = ai.GetContents();
                            foreach (GridColumn gc in q.GetColumns())
                            {
                                if (contents.ContainsKey(gc.GetName()))
                                {
                                    cellValues.Add(contents[gc.GetName()]);
                                }
                            } //end of foreach GridColumn
                            dg.Rows.Add(cellValues.ToArray<object>());
                        } //end of foreach AnswerItem
                    }
                    else if (q.GetQuestionType().ToLower() == "textboxdropdown")
                    {
                        Control ctl = this.allQControls[ans.GetName()];
                        //Question question = (Question)ctl.Tag;
                        List<string> strList= new List<string> ();
                        foreach (AnswerItem ai in ans.GetAnswerItems())
                        {
                            strList.Add(ai.GetValueName());
                        }                             
                            ctl.Text = string.Join(",", strList.ToArray());
                                              
                    
                    }
                    else  //for combobox and textbox
                    {
                        if (this.allQControls.ContainsKey(ans.GetName()))
                        {
                            Control ctl = this.allQControls[ans.GetName()];
                            //Question question = (Question)ctl.Tag;
                            ctl.Text = ans.GetValue();
                        }
                    }

                }   //end of "q is not null"
            }
        }

        public List<Answer> GetAnswer()
        {
            Dictionary<string, Answer> result = new Dictionary<string, Answer>();

            //找出所有 tag 值為 Question type 的控制項，
            foreach (Control ctrl in this.allQControls.Values)
            {
                Question q = ctrl.Tag as Question;
                if (q != null)
                {
                    if (!result.ContainsKey(q.GetQuestionName()))
                        result.Add(q.GetQuestionName(), new Answer());
                    Answer ans = result[q.GetQuestionName()];
                    ans.SetName(q.GetQuestionName());

                    if (ctrl is CheckBox)   //如果是 checkbox :
                    {
                        CheckBox chk = (CheckBox)ctrl;

                        if (chk.Checked)
                        {
                            //這是複選題
                            AnswerItem theAi = new AnswerItem();
                            string ctrlName = ctrl.Name;
                            QuestionListItem questionListItem = q.GetListItemByLable(ctrlName);
                            theAi.SetValueName(ctrlName);
                            if (questionListItem != null)
                            {
                                if (questionListItem.HasText)  //有 remark
                                {
                                    string remarkCtrlName = ctrlName + "_remark";
                                    if (this.allQControls.ContainsKey(remarkCtrlName))
                                    {
                                        theAi.SetValueRemark(this.allQControls[remarkCtrlName].Text);
                                    }
                                    else
                                        theAi.SetValueRemark("");
                                }
                            }
                            ans.AddAnswerItem(theAi);
                        } //end of (chk.Checked)
                    }
                    else if (ctrl is DataGridView)
                    {
                        DataGridView dg = (DataGridView)ctrl;
                        Dictionary<string, string> value = new Dictionary<string, string>();
                        foreach (DataGridViewRow row in dg.Rows)
                        {
                            if (!row.IsNewRow)
                            {
                                AnswerItem ai = new AnswerItem();
                                Dictionary<string, string> itemContents = new Dictionary<string, string>();
                                foreach (GridColumn gc in q.GetColumns())
                                {
                                    itemContents.Add(gc.GetName(), (row.Cells[gc.GetName()].Value == null) ? "" : row.Cells[gc.GetName()].Value.ToString());
                                }
                                ai.SetContent(itemContents);
                                ans.GetAnswerItems().Add(ai);
                            }
                        }
                    }
                    else   //視為單選題
                    {
                        if (q.GetQuestionType().ToLower() == "textboxdropdown")
                        {
                            ctrl.Text = ctrl.Text.Trim();
                            string[] strArr = ctrl.Text.Split(',');                         
                            foreach (string str in strArr)
                            {
                                if (!string.IsNullOrEmpty(str))
                                {
                                    XmlElement elm = new XmlDocument().CreateElement("Item");
                                    elm.SetAttribute("value", str.Trim ());
                                    AnswerItem ai = new AnswerItem(elm);
                                    ans.AddAnswerItem(ai);
                                }                           
                            }
                            
                            
                        }
                        else
                        ans.SetValue(ctrl.Text);
                    }
                }
            }


            return result.Values.ToList<Answer>();
        }

        public void ResetContent()
        {
            foreach (Control ctrl in this.allQControls.Values)
            {
                if (ctrl is CheckBox)   //如果是 checkbox :
                {
                    CheckBox chk = (CheckBox)ctrl;

                    if (chk.Checked)
                    {
                        chk.Checked = false;
                    }
                }
                else if (ctrl is DataGridView)
                {
                    DataGridView dg = (DataGridView)ctrl;
                    dg.Rows.Clear();
                    this.FillGridDefaultRecord(dg);
                }
                else   //視為單選題
                {
                    ctrl.Text = "";
                }
            }
        }

        private void makeUI()
        {
            List<QuestionListItem> listItems = this.questionGroup.GetListItems();
            this.allQControls.Clear();
            int totalLength = 0;
            int totalHeight = 0;
            foreach (Question q in this.questionGroup.GetQuestions())
            {
                Control ctrl = this.makeQuestionControl(q);
                this.contentPanel.Controls.Add(ctrl);

                if (totalHeight == 0)
                    totalHeight = ctrl.Height + 6;

                totalLength += ctrl.Width + 6;
                if (totalLength > this.contentPanel.Width)
                {
                    totalHeight += (ctrl.Height + 6);
                    totalLength = ctrl.Width + 6;
                }

                //totalVolumn += (ctrl.Width + 6) * (ctrl.Height + 6);
                /*
               if (q.GetQuestionType().ToLower() == "grid" )
               {
                   totalLength += (this.contentPanel.Width * (ctrl.Height / 25 - 1));
                   this.contentPanel.Height = 28 * (totalLength / this.contentPanel.Width + 1);                                        
               }
               else if (q.GetQuestionType().ToLower() == "checkbox")
               {
                   this.contentPanel.Height = ctrl.Height + 6;                                        
               }
               else if (q.GetQuestionType().ToLower() == "textarea")
               {
                   this.contentPanel.Height = ctrl.Height + 6;
               }
               else if (q.GetQuestionType().ToLower() == "text")
               {
                   totalLength += ctrl.Width + 10;
                   this.contentPanel.Height = 35 * (totalLength / this.contentPanel.Width + 1);
               }
               else if (q.GetQuestionType().ToLower() == "textboxdropdown")
               {
                   totalLength += ctrl.Width + 10;
                   this.contentPanel.Height = 22 * (totalLength / this.contentPanel.Width + 1);
               }
               else if (q.GetQuestionType().ToLower() == "combobox")
               {
                   totalLength += ctrl.Width + 10;
                   this.contentPanel.Height = 35 * (totalLength / this.contentPanel.Width + 1);
               }
               else 
               {
                   totalLength += ctrl.Width + 10;
                   this.contentPanel.Height = 25 * (totalLength / this.contentPanel.Width + 1);
               }
               */

                //int theHeight = totalVolumn / this.contentPanel.Width;
                //if (theHeight % 30 != 0)
                //    theHeight = 30 * (theHeight / 30 + 1);

                this.contentPanel.Height = totalHeight;

                this.pnlQGroup.Height = this.contentPanel.Height + 6;
                this.qType = q.GetQuestionType();   //假設同一題組裡的題型都是相同的。
            }
        }

        private void adjHeight()
        {
            int unitHeight = 20;
            int resultHeight = unitHeight * 2;
            if (this.questionGroup.GetQuestions().Count > 0)
            {
                Question q = this.questionGroup.GetQuestions()[0];
                if (q.GetQuestionType().ToLower() == "combobox")
                {
                    resultHeight = unitHeight * ((this.questionGroup.GetQuestions().Count / 5) + 2);
                }
            }
            this.pnlQGroup.Height = resultHeight;
        }

        //Factory Method
        private Control makeQuestionControl(Question q)
        {
            Control result = null;
            try
            {
                string qType = q.GetQuestionType();

                switch (qType.ToLower())
                {
                    case "combobox":
                        result = createComboBox(q);
                        break;
                    case "checkbox":
                        result = createCheckedListBox(q);
                        break;
                    case "text":
                        result = createTextBox(q);
                        break;
                    case "textarea":
                        result = createTextArea(q);
                        break;
                    case "grid":
                        result = createGrid(q);
                        break;
                    case "textboxdropdown":
                        result = createTextBoxDropDown(q);
                        break;
                    default:
                        result = createTextBox(q);
                        break;
                }
            }
            catch (Exception ex)
            {
                FISCA.Presentation.Controls.MsgBox.Show(ex.Message);                
            }
            return result;
        }

        private Control createTextBox(Question q)
        {
            int innerPanelLength = 0;
            FlowLayoutPanel pnlInner = new FlowLayoutPanel();
            pnlInner.FlowDirection = FlowDirection.LeftToRight;
            pnlInner.Width = 120;   //the default width

            //pnl.BackColor = System.Drawing.Color.Green;

            Label lbl = this.createLabel(q.GetQuestionLabel());
            pnlInner.Controls.Add(lbl);
            innerPanelLength += lbl.Width + 6;

            TextBox txt = new TextBox();
            txt.Enabled = false;
            txt.Name = q.GetQuestionName();
            //txt.Width = pnlInner.Width - lbl.Width - 20;
            txt.Height = 20;
            txt.TextAlign = HorizontalAlignment.Center;
            txt.Tag = q;
            txt.TextChanged += new EventHandler(txt_TextChanged);

            //調整寬度
            pnlInner.Width = innerPanelLength + txt.Width + 6;

            if (!string.IsNullOrEmpty(q.GetWidth()))
            {
                if (q.GetWidth().ToUpper() == "FILL")
                {
                    pnlInner.Width = this.contentPanel.Width - 12;
                    txt.Width = pnlInner.Width - innerPanelLength - 12;
                }
                else
                {
                    int width = 0;
                    bool isnum = int.TryParse(q.GetWidth(), out width);
                    if (isnum)
                    {
                        txt.Width = width;
                        pnlInner.Width = innerPanelLength + txt.Width + 6;
                    }
                }
            }

            //txt.Text = q.GetQuestionName();
            pnlInner.Controls.Add(txt);
            this.allQControls.Add(txt.Name, txt);

            //pnlInner.Width = innerPanelLength + txt.Width + 6;
            pnlInner.Height = txt.Height + 6;

            //adjust height
            int periodCount = 4;
            this.contentPanel.Height = 6 + pnlInner.Height * ((this.questionGroup.GetQuestions().Count % periodCount == 0) ? (this.questionGroup.GetQuestions().Count / periodCount + 1) : (this.questionGroup.GetListItems().Count / periodCount + 2));

            this.pnlQGroup.Height = this.contentPanel.Height + 6;
            return pnlInner;

        }

        private Control createTextArea(Question q)
        {
            int innerPanelLength = 0;
            FlowLayoutPanel pnlInner = new FlowLayoutPanel();
            pnlInner.FlowDirection = FlowDirection.LeftToRight;
            pnlInner.Width = 120;

            //pnl.BackColor = System.Drawing.Color.Green;            
            Label lbl = this.createLabel(q.GetQuestionLabel());
            pnlInner.Controls.Add(lbl);
            innerPanelLength += lbl.Width + 6;
            if (string.IsNullOrEmpty(q.GetQuestionLabel()))
                lbl.Height = 0;

            TextBox txt = new TextBox();
            txt.Enabled = false;
            txt.Name = q.GetQuestionName();
            txt.Multiline = true;
            txt.Width = pnlInner.Width - 6;
            txt.Height = q.GetRows() * 20;
            txt.TextAlign = HorizontalAlignment.Left;
            txt.ScrollBars = ScrollBars.Vertical;
            txt.Tag = q;
            txt.TextChanged += new EventHandler(txt_TextChanged);

            //調整寬度 , the default width of textarea  is FILL
            pnlInner.Width = this.contentPanel.Width - 6;
            txt.Width = pnlInner.Width - 6;

            if (!string.IsNullOrEmpty(q.GetWidth()))
            {
                if (q.GetWidth().ToUpper() != "FILL")
                {
                    int width = 0;
                    bool isnum = int.TryParse(q.GetWidth(), out width);
                    if (isnum)
                        txt.Width = width;
                }
            }

            //txt.Text = q.GetQuestionName();
            pnlInner.Controls.Add(txt);
            this.allQControls.Add(txt.Name, txt);

            //pnlInner.Width = this.contentPanel.Width - 6;
            pnlInner.Height = txt.Height + lbl.Height + 10;

            //adjust height
            this.contentPanel.Height = pnlInner.Height + 6;

            this.pnlQGroup.Height = this.contentPanel.Height + 6;
            return pnlInner;

        }

        private Control createTextBoxDropDown(Question q)
        {
            int innerPanelLength = 0;
            FlowLayoutPanel pnlInner = new FlowLayoutPanel();
            pnlInner.FlowDirection = FlowDirection.LeftToRight;
            pnlInner.Width = 120;
            //pnl.BackColor = System.Drawing.Color.Green;

            Label lbl = this.createLabel(q.GetQuestionLabel());
            pnlInner.Controls.Add(lbl);
            innerPanelLength += lbl.Width + 6;

            TextBoxDropDown txt = new TextBoxDropDown();
            txt.Enabled = false;
            txt.Name = q.GetQuestionName();
            txt.DropDownControl = this.getDropDownControl(q);
            txt.ButtonDropDown.Text = "...";
            txt.ButtonDropDown.Visible = true;
            txt.Width = 150;
            txt.Height = 20;
            txt.TextAlign = HorizontalAlignment.Left;
            txt.Tag = q;
            txt.TextChanged += new EventHandler(txt_TextChanged);
            txt.ButtonDropDownClick += new System.ComponentModel.CancelEventHandler(txt_ButtonDropDownClick);            

            //調整寬度
            pnlInner.Width = innerPanelLength + txt.Width + 6;

            if (!string.IsNullOrEmpty(q.GetWidth()))
            {
                if (q.GetWidth().ToUpper() == "FILL")
                {
                    pnlInner.Width = this.contentPanel.Width - 12;
                    txt.Width = pnlInner.Width - innerPanelLength - 12;
                }
                else
                {
                    int width = 0;
                    bool isnum = int.TryParse(q.GetWidth(), out width);
                    if (isnum)
                    {
                        txt.Width = width;
                        pnlInner.Width = innerPanelLength + txt.Width + 6;
                    }
                }
            }

            //txt.Text = q.GetQuestionName();
            pnlInner.Controls.Add(txt);
            this.allQControls.Add(txt.Name, txt);

            //pnlInner.Width = innerPanelLength + txt.Width + 6;
            pnlInner.Height = txt.Height + 6;

            //adjust height
            int periodCount = 4;
            this.contentPanel.Height = 6 + 20 * ((this.questionGroup.GetQuestions().Count % periodCount == 0) ? (this.questionGroup.GetQuestions().Count / periodCount + 1) : (this.questionGroup.GetListItems().Count / periodCount + 2));

            this.pnlQGroup.Height = this.contentPanel.Height + 6;
            return pnlInner;
        }

        private Control createComboBox(Question q)
        {
            FlowLayoutPanel pnlInner = new FlowLayoutPanel();
            pnlInner.FlowDirection = FlowDirection.LeftToRight;

            int innerPanelLength = 0;

            //如果 Question 有 Label ，則要在前面加上 Label
            if (q.HasLabel)
            {
                Label lbl = this.createLabel(q.GetQuestionLabel());
                pnlInner.Controls.Add(lbl);
                innerPanelLength += lbl.Width + 6;
            }

            //create combobox
            ComboBoxEx cbo = new ComboBoxEx();
            cbo.Enabled = false;
            cbo.Name = q.GetQuestionName();
            cbo.DropDownStyle = ComboBoxStyle.DropDown;
            cbo.Tag = q;
            cbo.TextChanged += new EventHandler(txt_TextChanged); ;
            Graphics g = cbo.CreateGraphics();
            SizeF maxSize = new SizeF();

            Font f = this.pnlQGroup.Font;
            foreach (QuestionListItem item in this.questionGroup.GetListItems())
            {
                int index = cbo.Items.Add(item.GetLabel());
                if (item.Selected)
                    cbo.SelectedText = item.GetLabel();

                SizeF theSize = g.MeasureString(item.GetLabel(), f);
                if (theSize.Width > maxSize.Width)
                    maxSize = theSize;
            }

            cbo.Width = (int)maxSize.Width + 25;
            pnlInner.Controls.Add(cbo);
            this.allQControls.Add(cbo.Name, cbo);

            pnlInner.Width = innerPanelLength + cbo.Width + 6;
            pnlInner.Height = cbo.Height + 6;

            //adjust height
            int periodCount = 4;
            this.contentPanel.Height = 6 + 20 * ((this.questionGroup.GetQuestions().Count % periodCount == 0) ? (this.questionGroup.GetQuestions().Count / periodCount + 1) : (this.questionGroup.GetQuestions().Count / periodCount + 2));

            this.pnlQGroup.Height = this.contentPanel.Height + 6;
            return pnlInner;
        }

        private Control createCheckedListBox(Question q)
        {
            FlowLayoutPanel pnl = new FlowLayoutPanel();
            pnl.FlowDirection = FlowDirection.LeftToRight;
            pnl.Width = this.contentPanel.Width - 6;
            pnl.WrapContents = true;

            int totalLength = 0;

            foreach (QuestionListItem item in this.questionGroup.GetListItems())
            {
                FlowLayoutPanel pnlInner = new FlowLayoutPanel();
                pnlInner.FlowDirection = FlowDirection.LeftToRight;

                CheckBox chk = new CheckBox();
                chk.Enabled = false;
                chk.Name = item.GetLabel();
                chk.Width = 15;
                chk.Height = 20;
                chk.Tag = q;
                chk.CheckedChanged += new EventHandler(txt_TextChanged);
                System.Windows.Forms.Padding margin = chk.Margin;
                margin.Top = 3;
                margin.Bottom = 0;
                margin.Right = 0;
                chk.Margin = margin;
                pnlInner.Controls.Add(chk);
                this.allQControls.Add(chk.Name, chk);

                Label lbl = this.createLabel(item.GetLabel());
                pnlInner.Controls.Add(lbl);

                pnlInner.Width = chk.Width + lbl.Width + 25;
                pnlInner.Height = chk.Height + 2;

                if (item.HasText)
                {
                    TextBox txt = new TextBox();
                    txt.Name = chk.Name + "_remark";
                    txt.Width = 80;
                    txt.TextChanged += new EventHandler(txt_TextChanged); ;
                    //txt.Multiline = true;
                    //txt.Height = 26;
                    //txt.Tag = q;  //此 TextBox 請不要設定Tag，只要加入 allQControls 就好。
                    pnlInner.Controls.Add(txt);
                    this.allQControls.Add(txt.Name, txt);
                    pnlInner.Width += txt.Width;
                }

                pnl.Controls.Add(pnlInner);
                totalLength += pnlInner.Width + 10;

            }
            //adjust height
            //int periodCount = 4;
            //pnl.Height = 22 * ((this.questionGroup.GetListItems().Count % periodCount == 0) ? (this.questionGroup.GetListItems().Count / periodCount + 1) : (this.questionGroup.GetListItems().Count / periodCount + 2));
            pnl.Height = 28 * (totalLength / pnl.Width + 1);
            //this.contentPanel.Height = pnl.Height + 6;
            //this.pnlQGroup.Height = this.contentPanel.Height + 6;

            return pnl;
        }

        private Control createGrid(Question q)
        {
            FlowLayoutPanel pnl = new FlowLayoutPanel();
            pnl.FlowDirection = FlowDirection.LeftToRight;
            pnl.Width = this.contentPanel.Width - 6;
            pnl.WrapContents = true;

            DataGridViewX dg = new DataGridViewX();
            dg.Enabled = false;
            dg.Tag = q;
            dg.ColumnCount = q.GetColumns().Count;
            dg.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;

            // Set the column header style.
            DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle();
            columnHeaderStyle.BackColor = Color.Aqua;
            dg.BackgroundColor = Color.White;            
            columnHeaderStyle.Font = new Font("微軟粗黑體", 9, FontStyle.Regular);
            dg.ColumnHeadersDefaultCellStyle = columnHeaderStyle;

            for (int i = 0; i < q.GetColumns().Count; i++)
            {
                GridColumn gc = q.GetColumns()[i];
                dg.Columns[i].Name = gc.GetName();
                if (!string.IsNullOrEmpty(gc.GetWidth()))
                    dg.Columns[i].Width = int.Parse(gc.GetWidth());
            }

            dg.Width = pnl.Width - 10;
            dg.Height = 120;
            FillGridDefaultRecord(dg);

            dg.CellEndEdit += new DataGridViewCellEventHandler(dg_CellEndEdit);
            pnl.Height = dg.Height + 6;
            this.allQControls.Add(q.GetQuestionName(), dg);
            pnl.Controls.Add(dg);

            this.contentPanel.Height = pnl.Height + 6;
            this.pnlQGroup.Height = this.contentPanel.Height + 6;
            return pnl;
        }

        private void FillGridDefaultRecord(DataGridView dg)
        {
            Question q = (Question)dg.Tag;

            //fill the default record !
            if (q.GetDefaultRecords().Count > 0)
            {
                foreach (Dictionary<string, string> key_values in q.GetDefaultRecords())    //for every record
                {
                    List<object> cellValues = new List<object>();
                    Dictionary<string, string> contents = key_values;
                    foreach (GridColumn gc in q.GetColumns())
                    {
                        if (contents.ContainsKey(gc.GetName()))
                            cellValues.Add(contents[gc.GetName()]);
                        else
                            cellValues.Add("");
                    } //end of foreach GridColumn
                    dg.Rows.Add(cellValues.ToArray<object>());
                }
            }
        }

        void txt_ButtonDropDownClick(object sender, System.ComponentModel.CancelEventArgs e)
        {
            this.textBoxDropDownDecorator.SetTextBox((TextBoxDropDown)sender);
        }

        //確認 TextBoxDropDown 的 DropDownControl 一定會有值
        private Control getDropDownControl(Question q)
        {
            if (this.chkBox == null)
            {
                this.chkBox = new CheckedListBox();
                this.chkBox.CheckOnClick = true;
                this.chkBox.Width = 150;
                this.chkBox.Height = 100;
                this.textBoxDropDownDecorator = new TextBoxDropDownDecorator(this.chkBox);
                foreach (QuestionListItem s in q.GetListItems())
                {
                    this.chkBox.Items.Add(s.GetLabel());
                }
            }
            return this.chkBox;
        }

        void txt_TextChanged(object sender, EventArgs e)
        {
            this.RaiseEvent(sender);
        }


        void dg_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            this.RaiseEvent(sender);
        }

        private Label createLabel(string text)
        {
            Label lbl = new Label();
            lbl.Text = text;
            lbl.Width = (int)lbl.CreateGraphics().MeasureString(text, this.pnlQGroup.Font).Width + 6;
            System.Windows.Forms.Padding margin = lbl.Margin;
            margin.Top = 3;
            margin.Right = 0;
            lbl.Margin = margin;

            return lbl;
        }

        private void RaiseEvent(object sender)
        {
            if (this.OnContentChange != null)
                this.OnContentChange(sender, new EventArgs());
        }
    }
}
