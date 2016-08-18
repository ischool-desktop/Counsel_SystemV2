using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace Counsel_System2.UI
{
    public partial class SubjectUIDefinition : UserControl
    {
        private XmlElement elmSubject;
        private string label = "";
        private Parser.Subject subj;
        private QuestionGroupUIDefinition qGroupUIDef;
        private QuestionUIDefinition qUIDef;

        public SubjectUIDefinition()
        {
            InitializeComponent();
        }

        private void SubjectUIDefinition_Load(object sender, EventArgs e)
        {
            
        }

        public void SetLabel(string label)
        {
            this.label = label;
        }

        public void SetSubjectDefinition(XmlElement elmSubject)
        {
            this.elmSubject = elmSubject;
            parseContent();
        }



        private void parseContent()
        {
            this.subj = new Parser.Subject(this.elmSubject);
            initTreeNode();
        }

        private void initTreeNode()
        {
            this.advTree1.Nodes.Clear();

            foreach (Parser.QuestionGroup qgroup in this.subj.GetQuestionGroups())
            {
                DevComponents.AdvTree.Node tn = new DevComponents.AdvTree.Node();
                tn.Text = qgroup.GetLabel();
                tn.Tag = qgroup;
                foreach (Parser.Question q in qgroup.GetQuestions())
                {
                    DevComponents.AdvTree.Node tnQ = new DevComponents.AdvTree.Node();
                    tnQ.Text = string.Format("{0} ({1})", q.GetQuestionLabel(), q.GetQuestionName());
                    tnQ.Tag = q;
                    tn.Nodes.Add(tnQ);
                }
                this.advTree1.Nodes.Add(tn);
            }
        }

        private void advTree1_NodeClick(object sender, DevComponents.AdvTree.TreeNodeMouseEventArgs e)
        {
            UpdateDef();

            object obj = e.Node.Tag;

            if (obj is Parser.QuestionGroup)
            {
                showQuestionGroupDefinition((Parser.QuestionGroup)obj);
                showPreview((Parser.QuestionGroup)obj);
            }
            else
            {
                showQuestionDefinition((Parser.Question)obj);
            }           
        }

        private void showPreview(Parser.QuestionGroup qgroup)
        {
            QGPanel pnl = new QGPanel();
            pnl.Width = this.pnlPreview.Width - 10;
            pnl.SetDefinition(qgroup);
            pnl.Top = 5;
            pnl.Left = 5;
            
            this.pnlPreview.Controls.Clear();
            this.pnlPreview.Controls.Add(pnl);
        }

        private void showQuestionGroupDefinition(Parser.QuestionGroup qgroup)
        {
            if (this.qGroupUIDef == null) 
                this.qGroupUIDef = new QuestionGroupUIDefinition();

            this.qGroupUIDef.SetQuestionGroup(qgroup);
            this.qGroupUIDef.Dock = DockStyle.Fill;

            this.pnlDefinition.Controls.Clear();
            this.pnlDefinition.Controls.Add(this.qGroupUIDef);

            
        }

        private void showQuestionDefinition(Parser.Question q)
        {
            if (this.qUIDef == null)
                this.qUIDef = new QuestionUIDefinition();

            this.qUIDef.SetQuestion(q);
            this.qUIDef.Dock = DockStyle.Fill;

            this.pnlDefinition.Controls.Clear();
            this.pnlDefinition.Controls.Add(this.qUIDef);

                
        }
        private void UpdateDef()
        {
            if (this.pnlDefinition.Controls.Count == 0)
                return;

            Control ctrl = this.pnlDefinition.Controls[0];
            if (ctrl is UI.QuestionGroupUIDefinition)
                ((QuestionGroupUIDefinition)ctrl).UpdateNewValue();
            else
                ((QuestionUIDefinition)ctrl).UpdateNewValue();
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            UpdateDef();

            object obj = this.advTree1.SelectedNode.Tag;

            showPreview((Parser.QuestionGroup)obj);

            /*
            if (obj is Parser.QuestionGroup)
            {
                QGPanel pnl = new QGPanel();
                pnl.Width = this.pnlPreview.Width - 10;
                pnl.SetDefinition((Parser.QuestionGroup)obj);
                pnl.Top = 5;
                pnl.Left = 5;

                this.pnlPreview.Controls.Clear();
                this.pnlPreview.Controls.Add(pnl);
            }
             * */

        }

    }
}
