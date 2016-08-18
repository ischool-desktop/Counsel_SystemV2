using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Xml;
using Counsel_System2.Parser;
using Counsel_System2.UI;

namespace Counsel_System2
{
    public partial class QGPanel : System.Windows.Forms.FlowLayoutPanel
    {
        // Declare the delegate (if using non-generic pattern).
        public delegate void ContentChange(object sender, EventArgs e);
        // Declare the event.
        public event ContentChange OnContentChange;

        private QuestionGroup titleParser;
        private QGUIMaker uim;

        public QGPanel()
        {
            InitializeComponent();
            this.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Font = new System.Drawing.Font("微軟粗黑體", 10);
        }

        public QGPanel(IContainer container)
        {
            container.Add(this);
            InitializeComponent();
        }

        public void ResetContent()
        {
            this.uim.ResetContent();
        }

        public void SetDefinition(QuestionGroup qGroup)
        {
            this.titleParser = qGroup;
            this.uim = new QGUIMaker(this, this.titleParser);
            this.uim.OnContentChange += new QGUIMaker.ContentChange(uim_OnContentChange);
            //this.BackColor = System.Drawing.Color.LightYellow;
        }

        void uim_OnContentChange(object sender, EventArgs e)
        {
            if (this.OnContentChange != null)
                this.OnContentChange(sender, e);
        }

        public string GetLabel()
        {
            return this.titleParser.GetLabel();
        }

        public string GetIdentifyLabel()
        {
            return this.titleParser.GetIdentifyLabel();
        }

        public void SetAnswer(XmlElement xmlAnswers)
        {
            this.uim.SetAnswer(xmlAnswers);
        }

        public List<Answer> GetAnswers()
        {
            return this.uim.GetAnswer();
        }
    }
}
