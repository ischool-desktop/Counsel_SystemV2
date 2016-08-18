using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace Counsel_System2.Parser
{
    public class QuestionListItem
    {
        private string label = "";
        private bool hasText = false;
        private bool selected = false;

        public QuestionListItem()
        {
        }

        public QuestionListItem(XmlElement item)
        {
            if (item != null)
            {
                this.label = item.GetAttribute("label");
                this.hasText = (item.GetAttribute("hasText").ToUpper() == "TRUE");
                this.selected = (item.GetAttribute("selected").ToUpper() == "TRUE");
            }
        }

        public string GetLabel()
        {
            return this.label;
        }
        public void SetLabel(string label)
        {
            this.label = label;
        }

        public bool HasText
        {
            get { return this.hasText; }
            set { this.hasText = value; }
        }


        public bool Selected
        {
            get { return this.selected; }
            set { this.selected = value; }
        }
    }
}
