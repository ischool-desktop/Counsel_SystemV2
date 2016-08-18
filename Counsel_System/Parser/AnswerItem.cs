using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace Counsel_System2.Parser
{
    /**
     * 代表複選題的其中一個答案，適用 checkedlistbox 及 grid
     * <Item value=”其它” remark=”蛀牙” />
     * */
    public class AnswerItem
    {
        private string valueName = "";    // for checkbox only
        private string valueRemark = null;    //for checkbox only
        private Dictionary<string, string> values;      //for grid value

        public AnswerItem()
        {
            this.values = new Dictionary<string, string>();
        }

        public AnswerItem(XmlElement elm)
        {
            values = new Dictionary<string, string>();
            /*
            foreach (XmlAttribute att in elm.Attributes)
            {
                if (!values.ContainsKey(att.Name))
                    values.Add(att.Name, att.Value);

                values[att.Name] = att.Value;
            }*/
            foreach (XmlElement elmField in elm.SelectNodes("Field"))
            {
                if (!values.ContainsKey(elmField.GetAttribute("key")))
                    values.Add(elmField.GetAttribute("key"), elmField.GetAttribute("value"));

                values[elmField.GetAttribute("key")] = elmField.GetAttribute("value");
            }

            this.valueName = elm.GetAttribute("value");
            if (elm.Attributes["remark"] != null)
                this.valueRemark = elm.GetAttribute("remark");
        }


        /// <summary>
        /// 設定 CheckBox 的 AnswerItem 的 value 屬性值，如 <Item value=”氣喘” />
        /// </summary>
        /// <param name="valueName"></param>
        public void SetValueName(string valueName)
        {
            this.valueName = valueName;
        }

        /// <summary>
        /// for checkbox only
        /// </summary>
        /// <returns></returns>
        public string GetValueName()
        {
            string result = this.valueName;
            //if (!string.IsNullOrEmpty(this.GetValueRemark()))
            //    result += "_remark";    //checkbox 某選項的 "其它" 文字方塊的命名規則

            return this.valueName;
        }


        /// <summary>
        /// 設定 CheckBox 的 AnswerItem 的 remark 屬性值，如 <Item value=”其它” remark=”蛀牙” />
        /// </summary>
        /// <returns></returns>
        public void SetValueRemark(string remarkValue)
        {
            this.valueRemark = remarkValue;
        }

        /// <summary>
        /// for checkbox only
        /// </summary>
        /// <returns></returns>
        public string GetValueRemark()
        {
            return this.valueRemark;
        }

        public bool hasRemark
        {
            get { return this.valueRemark != null; }
        }


        /// <summary>
        /// 取得一個 AnswerItem 裡某一項屬性的值，是給 Grid 傳入欄位名稱取得某一值用的
        /// </summary>
        /// <param name="attName"></param>
        /// <returns></returns>
        public string GetValue(string attName)
        {
            string result = "";
            if (this.values.ContainsKey(attName))
                result = this.values[attName];

            return result;
        }

        /// <summary>
        /// 取得 grid 的值專用
        /// </summary>
        /// <returns></returns>
        public Dictionary<string, string> GetContents()
        {
            return this.values;
        }

        /// <summary>
        /// 設定 grid 的值
        /// </summary>
        /// <returns></returns>
        public void SetContent(Dictionary<string, string> contents)
        {
            this.values = contents;
        }
    }
}
