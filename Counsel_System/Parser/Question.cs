using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace Counsel_System2.Parser
{
    /// <summary>
    /// 
    /// </summary>
    public class Question
    {
        private string subjectName = "";
        private List<QuestionListItem> listItems;
        private string qType = "";
        private string qName = "";
        private string qLabel = "";
        private string qWidth = "";
        private string qRows = "";  //assign the height of textarea control , or the rowCount of Grid Control !
        private List<GridColumn> columns;
        private List<Dictionary<string, string>> defaultRecords;    //the default record of grid

        /**
			    <Question type="checkedlistbox" label="血型" />
         * */
        public Question(string subjectName, List<QuestionListItem> listItems, XmlElement question)
        {
            this.subjectName = subjectName;
            this.listItems = listItems;
            this.qType = question.GetAttribute("type");
            this.qName = question.GetAttribute("name");
            this.qLabel = question.GetAttribute("label");
            this.qWidth = question.GetAttribute("width");
            this.qRows = question.GetAttribute("rows");
            if (string.IsNullOrEmpty(this.qRows))
                this.qRows = "2";   //Default value;

            //parse column definitions
            this.columns = new List<GridColumn>();
            foreach (XmlElement xmlCol in question.SelectNodes("Cols/Col"))
            {
                this.columns.Add(new GridColumn(xmlCol));
            }

            //parse default values
            this.defaultRecords = new List<Dictionary<string, string>>();

            foreach (XmlElement xmlItem in question.SelectNodes("Default/Item"))  //every record
            {
                Dictionary<string, string> key_values = new Dictionary<string, string>();
                this.defaultRecords.Add(key_values);
                foreach (XmlElement elmField in xmlItem.SelectNodes("Field"))
                {
                    key_values.Add(elmField.GetAttribute("key"), elmField.GetAttribute("value"));
                }
            }
        }

        public string GetQuestionType()
        {
            return this.qType;
        }
        public string GetQuestionName()
        {
            return this.qName;
        }
        public string GetQuestionLabel()
        {
            return this.qLabel;
        }

        public bool HasLabel
        {
            get { return !string.IsNullOrEmpty(this.qLabel); }
        }

        public string GetQuestionFullName()
        {
            StringBuilder sb = new StringBuilder(this.subjectName);

            if (!string.IsNullOrEmpty(this.qName))
                sb.Append("_");

            sb.Append(this.qName);

            return sb.ToString();
        }

        public string GetWidth()
        {
            return this.qWidth;
        }

        public void SetWidth(string width)
        {
            this.qWidth = width;
        }

        public int GetRows()
        {
            return int.Parse(this.qRows);
        }

        public void SetRows(int rows)
        {
            this.qRows = rows.ToString();
        }

        public List<QuestionListItem> GetListItems()
        {
            return this.listItems;
        }

        public QuestionListItem GetListItemByLable(string label)
        {
            QuestionListItem result = null;
            foreach (QuestionListItem it in this.listItems)
            {
                if (it.GetLabel().ToUpper() == label.ToUpper())
                {
                    result = it;
                    break;
                }
            }

            return result;
        }

        public List<GridColumn> GetColumns()
        {
            return this.columns;
        }

        public List<Dictionary<string, string>> GetDefaultRecords()
        {
            return this.defaultRecords;
        }
    }
}
