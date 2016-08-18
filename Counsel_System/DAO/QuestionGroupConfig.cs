using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Counsel_System2.DAO
{

    /// <summary>
    /// 綜合表現樣板 Question 使用(問題群組)
    /// </summary>
    public class QuestionGroupConfig
    {
        XElement _QGroup;
        /// <summary>
        ///Question Type
        /// </summary>
        public enum QuestionType{text,combobox,grid,checkbox,textboxdropdown}
        
        public QuestionGroupConfig(XElement QGroup)
        {
            if (QGroup == null)
            {
                _QGroup = new XElement("QG");

            }
            _QGroup = QGroup;
        }

        /// <summary>
        /// 設定問題群組標籤
        /// </summary>
        /// <param name="label"></param>
        public void SetQuestionGroupLabel(string label)
        {
            _QGroup.SetAttributeValue("label", label);
        }


        /// <summary>
        /// Add Question Group Choices
        /// </summary>
        /// <param name="label"></param>
        /// <param name="selected"></param>
        public void AddChoice(string label,bool selected)
        { 
            if(_QGroup.Element("Choice") == null)
                _QGroup.SetElementValue("Choice","");

            XElement elm = new XElement ("Item");
            elm.SetAttributeValue("label",label );
            if(selected)
                elm.SetAttributeValue("selected","true");
            
            _QGroup.Element("Choice").Add(elm);        
        }

        /// <summary>
        /// Get Question Group Choices XElement
        /// </summary>
        /// <returns></returns>
        public XElement GetChoiceXml()
        { 
            XElement retVal = new XElement ("Choices");
            
            if(_QGroup.Element("Choice") != null )
                retVal=_QGroup.Element("Choice");

            return retVal ;
        }

        public void SetChoiceXml(XElement ChoiceXml)
        {
            if (_QGroup.Element("Choice") != null)
                _QGroup.Element("Choice").Remove();

                _QGroup.Add(ChoiceXml);        
        }

        /// <summary>
        /// 新增選擇問題
        /// </summary>
        /// <param name="Qtype"></param>
        /// <param name="name"></param>
        /// <param name="label"></param>
        /// <param name="width"></param>
        /// <param name="rows"></param>
        /// <param name="GridColXml"></param>
        /// <param name="GridDefaultXml"></param>
        public void AddQuestionSelect(QuestionType Qtype, string name, string label, string width, int rows,XElement  GridColXml,XElement GridDefaultXml)
        {
            if (_QGroup.Element("Qs") == null)
                _QGroup.SetElementValue("Qs", "");

            XElement elm = new XElement("Q");
            elm.SetAttributeValue("type", Qtype.ToString());
            elm.SetAttributeValue("name", name);
            elm.SetAttributeValue("label", label);
            elm.SetAttributeValue("width", width);

            // 解析使用 Grid 型態
            if (Qtype == QuestionType.grid)
            {
                elm.SetAttributeValue("rows", rows.ToString());
                elm.Add(GridColXml);
                elm.Add(GridDefaultXml);
            }            
        }

        /// <summary>
        /// 取得所有問題
        /// </summary>
        /// <returns></returns>
        public XElement GetQuestionSelectXml()
        {
            XElement retVal = new XElement("Qs");
            if (_QGroup.Element("Qs") != null)
                retVal = _QGroup.Element("Qs");

            return retVal;
        }     

    }
}
