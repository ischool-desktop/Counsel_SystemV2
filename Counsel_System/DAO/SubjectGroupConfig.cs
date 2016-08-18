using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// 綜合表現使用,資料存取 XML 轉換
    /// </summary>
    public class SubjectGroupConfig
    {
        XElement _Template;

        /// <summary>
        /// 主題名稱
        /// </summary>
        public string SubjectName { get; set; }

        /// <summary>
        /// 載入預設樣版
        /// </summary>
        /// <param name="template"></param>
        public SubjectGroupConfig(string TemplateStr)
        {
            StringBuilder sb = new StringBuilder ();
            sb.Append("<Root>");sb.Append(TemplateStr);sb.Append("</Root>");

            _Template = XElement.Parse(sb.ToString());
        }

        /// <summary>
        /// 取得主題內label
        /// </summary>
        /// <returns></returns>
        public string GetSubjectLabel()
        {
            string retVal = "";

            if (_Template.Element("Subject") != null)
                if (_Template.Element("Subject").Attribute("label") != null)
                    retVal = _Template.Element("Subject").Attribute("label").Value;

            return retVal;
        }

        /// <summary>
        /// 設定主題內 label 名稱
        /// </summary>
        /// <param name="label"></param>
        public void SetSubjectLabel(string label)
        {
            if (_Template.Element("Subject") != null)
                _Template.Element("Subject").SetAttributeValue("label", label);
        
        }





        /// <summary>
        /// 取得儲存 XML 字串結構，回存資料庫使用
        /// </summary>
        /// <returns></returns>
        public string GetTemplateXMLString()
        {
            string retVal = "";
            if (_Template.Element("Subject") != null)
                retVal = _Template.Element("Subject").Value;

            return retVal;
        }
    }
}
