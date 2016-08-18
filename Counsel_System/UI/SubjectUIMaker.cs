using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Windows.Forms;
using System.Drawing;
using System.Xml.Linq;

namespace Counsel_System2.UI
{
    /// <summary>
    /// Render 整個 大題的 UI
    /// </summary>
    public class SubjectUIMaker
    {
        // Declare the delegate (if using non-generic pattern).
        public delegate void ContentChange(object sender, EventArgs e);
        // Declare the event.
        public event ContentChange OnContentChange;

        private string subjectLabel = "";
        private XmlElement xmlSubjectDef;
        private XmlElement xmlAnswers;
        private Panel pnlContainer;
        private Dictionary<string, QGPanel> questionGroupPanels;
        private Parser.Subject subj;

        private int y_pos = 0;
        private int padding = 5;

        public SubjectUIMaker(XmlElement xmlSubjectDef, System.Windows.Forms.Panel pnlContainer)
        {
            this.subjectLabel = xmlSubjectDef.GetAttribute("label");
            this.xmlSubjectDef = xmlSubjectDef;
            this.pnlContainer = pnlContainer;
            this.pnlContainer.Controls.Clear();
            makeUI();
        }

        public void SetAnswer(XmlElement answers)
        {
            this.xmlAnswers = answers;
            this.ResetContent();
            foreach (QGPanel pnl in this.questionGroupPanels.Values)
            {
                pnl.SetAnswer(this.xmlAnswers);
            }
        }

        public String GetAnswers()
        {
            XElement elmAnswers = new XElement("Answers");
            elmAnswers.SetAttributeValue("label", this.subjectLabel);
            foreach (QGPanel pnl in this.questionGroupPanels.Values)
            {
                foreach (Parser.Answer ans in pnl.GetAnswers())
                    elmAnswers.Add(ans.GetXml());
            }
            return elmAnswers.ToString();


            //StringBuilder sb = new StringBuilder(string.Format("<Answers label='{0}'>", this.subjectLabel));

            //foreach (QGPanel pnl in this.questionGroupPanels.Values)
            //{
            //    foreach (Parser.Answer ans in pnl.GetAnswers())
            //    {
            //        sb.Append(ans.GetXmlString());
            //    }
            //}

            //sb.Append("</Answers>");
            //return sb.ToString();
        }

        public void ResetContent()
        {
            foreach (QGPanel pnl in this.questionGroupPanels.Values)
            {
                pnl.ResetContent();
            }
        }

        private void makeUI()
        {
            this.subj = new Parser.Subject(this.xmlSubjectDef);

            this.questionGroupPanels = new Dictionary<string, QGPanel>();

            //foreach (XmlElement elm in this.xmlSubjectDef.SelectNodes("QG"))
            //{
            foreach (Parser.QuestionGroup qGroup in this.subj.GetQuestionGroups())
            {
                QGPanel pn = new QGPanel();
                pn.Width = this.pnlContainer.Width - 10;
                pn.SetDefinition(qGroup);
                //pn.SetDefinition(elm);
                //pn.SetAnswer(this.xmlAnswers);
                //this.flowLayoutPanel1.Controls.Add(pn);
                Point p = new Point(this.padding, this.y_pos + this.padding);
                pn.Location = p;

                pn.OnContentChange += new QGPanel.ContentChange(pn_OnContentChange);

                this.pnlContainer.Controls.Add(pn);
                this.y_pos = pn.Location.Y + pn.Height;

                this.questionGroupPanels.Add(pn.GetIdentifyLabel(), pn);
            }

            this.pnlContainer.Height = this.y_pos + 5;
        }

        void pn_OnContentChange(object sender, EventArgs e)
        {
            if (this.OnContentChange != null)
            {
                this.OnContentChange(sender, e);
            }
        }
    }
}
