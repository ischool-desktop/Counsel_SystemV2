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
    public partial class QuestionUIDefinition : UserControl
    {
        public QuestionUIDefinition()
        {
            InitializeComponent();
        }

        public void SetQuestion(Parser.Question q)
        {
            this.comboBoxEx1.SelectedItem = q.GetQuestionType();
        }

        public void UpdateNewValue()
        {
        }
    }
}
