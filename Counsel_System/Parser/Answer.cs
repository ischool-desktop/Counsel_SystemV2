using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace Counsel_System2.Parser
{
    /// <summary>
    /// 代表一個題目的答案。
    /// </summary>    
    public class Answer
    {
        /**
     * <Ans name="AAA12345677" value="61"></Ans>
	    <Ans name="AAA12345678" >
		    <Item value="近視" />
		    <Item value="其它" remark="蛀牙" />
	    </Ans>
     * */

        private string qname = "";  //題目名稱，應該是 Question.GetQuestionName()
        private string qvalue = "";     //for 單選題

        private List<AnswerItem> ansItems;      //for 複選題，如 checkbox, grid

        public Answer()
        {
            this.ansItems = new List<AnswerItem>();
        }

        public Answer(XmlElement elmAns)
        {
            this.qname = elmAns.GetAttribute("name");
            this.qvalue = elmAns.GetAttribute("value");
            ansItems = new List<AnswerItem>();
            foreach (XmlElement elmItem in elmAns.SelectNodes("Item"))
            {
                AnswerItem ai = new AnswerItem(elmItem);
                ansItems.Add(ai);
            }
        }

        public void SetName(string qname)
        {
            this.qname = qname;
        }

        public string GetName()
        {
            return this.qname;
        }

        /// <summary>
        /// 設定單選題的答案
        /// </summary>
        /// <param name="qvalue"></param>
        public void SetValue(string qvalue)
        {
            this.qvalue = qvalue;
        }

        /// <summary>
        /// 取得單選題的答案
        /// </summary>
        /// <param name="qvalue"></param>
        public string GetValue()
        {
            return this.qvalue;
        }

        public bool IsGrid { get; set; }


        /// <summary>
        /// 設定複選題答案，適用 checkedlistbox 及 grid
        /// </summary>
        /// <param name="answerItems"></param>
        public void SetAnswerItems(List<AnswerItem> answerItems)
        {
            this.ansItems = answerItems;
        }

        /// <summary>
        /// 取得複選題答案，適用 checkedlistbox 及 grid
        /// </summary>
        /// <returns></returns>
        public List<AnswerItem> GetAnswerItems()
        {
            return this.ansItems;
        }

        public void AddAnswerItem(AnswerItem ansItem)
        {
            bool isExisted = false;

            for (int i = 0; i < this.ansItems.Count; i++)
            {
                if (this.ansItems[i].GetValueName().ToUpper() == ansItem.GetValueName().ToUpper())
                {
                    this.ansItems[i] = ansItem;
                    isExisted = true;
                    break;
                }
            }

            if (!isExisted)
                this.ansItems.Add(ansItem);
        }

        public XElement GetXml()
        {
            XElement elmAns = new XElement("Ans");
            elmAns.SetAttributeValue("name", this.qname);
            elmAns.SetAttributeValue("value", this.qvalue);

            foreach (AnswerItem ai in this.ansItems)
            {
                
                if (!string.IsNullOrEmpty(ai.GetValueName()))   //checkbox
                {
                    XElement elmItem = new XElement("Item");
                    elmItem.SetAttributeValue("value", ai.GetValueName());
                    if (ai.GetValueRemark() != null)
                        elmItem.SetAttributeValue("remark", ai.GetValueRemark());
                    elmAns.Add(elmItem);
                }
                else   //grid
                {
                    XElement elmItem = new XElement("Item");
                    Dictionary<string, string> contents = ai.GetContents();
                    foreach (string key in contents.Keys)
                    {
                        XElement elmField = new XElement("Field");
                        elmField.SetAttributeValue("key", key);
                        elmField.SetAttributeValue("value", contents[key]);
                        elmItem.Add(elmField);
                    }
                    elmAns.Add(elmItem);
                }
                
            }
            return elmAns;
        }

        public string  GetXmlString()
        {
            StringBuilder sb = new StringBuilder(string.Format("<Ans name='{0}' value='{1}'>", this.qname, this.qvalue));
            foreach (AnswerItem ai in this.ansItems)
            {
                if (!string.IsNullOrEmpty(ai.GetValueName()))   //checkbox
                {
                    sb.Append(string.Format("<Item value='{0}' ", ai.GetValueName()));
                    if (ai.GetValueRemark() != null)
                        sb.Append(string.Format(" remark='{0}' ", ai.GetValueRemark()));
                    sb.Append(" />");
                }
                else   //grid
                {
                    sb.Append("<Item >");
                    Dictionary<string, string> contents = ai.GetContents();
                    foreach (string key in contents.Keys)
                    {
                        sb.Append(string.Format("<Field  key='{0}' value='{1}' />", key, contents[key]));
                    }
                    sb.Append(" </Item>");
                }
            }
            sb.Append("</Ans>");
            return sb.ToString();
        }
    }
}
