using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace Counsel_System2.Parser
{
    public class QuestionGroup
    {
        private XmlElement elmQGDef;
        private string label = "";
        private string tempLabel = "";  //如果label = "" 時，會造成 SubjectUIMaker 中的 Dictionary 會重覆""
        private string labelWidth = "";
        private bool showLabel = false;
        private List<QuestionListItem> listItems;
        private List<Question> questions;

        public QuestionGroup(XmlElement elmQGDef)
        {
            this.elmQGDef = elmQGDef;

            //parse titleName
            this.label = elmQGDef.GetAttribute("label");

            this.labelWidth = elmQGDef.GetAttribute("width");

            //如果label = "" 時，可能會造成 Subject 中的 Dictionary 會重覆""，所以先給他一個 Guid 值
            if (string.IsNullOrEmpty(this.label))
                this.tempLabel = Guid.NewGuid().ToString();
            else
                this.tempLabel = this.label;

            //parse hideLabel attribute
            this.showLabel = (elmQGDef.GetAttribute("hideLabel").ToUpper() != "TRUE");

            //parse ListItems
            this.listItems = new List<QuestionListItem>();
            foreach (XmlElement elm in elmQGDef.SelectNodes("Choices/Item"))
            {
                this.listItems.Add(new QuestionListItem(elm));
            }

            //Questions
            this.questions = new List<Question>();
            foreach (XmlElement elm in elmQGDef.SelectNodes("Qs/Q"))
            {
                Question q = new Question(this.label, this.listItems, elm);
                this.questions.Add(q);
            }
        }

        public XmlElement GetQGDefinition()
        {
            return this.elmQGDef;
        }

        public string GetLabel()
        {
            return this.label;
        }

        public string GetIdentifyLabel()
        {
            return this.tempLabel;
        }

        public void SetLabel(string label)
        {
            this.label = label;
        }

        public String GetLabelWidth()
        {
            return this.labelWidth;
        }

        public List<QuestionListItem> GetListItems()
        {
            return this.listItems;
        }

        public List<Question> GetQuestions()
        {
            return this.questions;
        }

        public bool GetLabelVisible()
        {
            return this.showLabel;
        }
        public void SetLabelVisible(bool labelVisible)
        {
            this.showLabel = !labelVisible;
        }

        public Question GetQuestionByName(string qName)
        {
            Question result = null;
            foreach (Question q in this.questions)
            {
                if (q.GetQuestionName().ToUpper() == qName.ToUpper())
                {
                    result = q;
                    break;
                }
            }

            return result;
        }
    }
}
