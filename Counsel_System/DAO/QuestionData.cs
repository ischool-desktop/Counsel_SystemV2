using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// 題目資料
    /// </summary>
    public class QuestionData
    {

        UDTQuestionsDataDef _update = null;

        public QuestionData(UDTQuestionsDataDef data)
        {
            if (data == null)
                _update = new UDTQuestionsDataDef();
            else
                _update = data;
        }

        public string Name { get { return _update.Name; } set { _update.Name = value; } }

        public string Group { get { return _update.Group; } set { _update.Group = value; } }

        /// <summary>
        /// 控制項類型
        /// </summary>
        public EnumControlType QControlType
        {
            get 
            {
                string str = _update.ControlType;

                if (_update.ControlType == "CHECKBOX")
                    return EnumControlType.CHECKBOX;
                else if (_update.ControlType == "COMBOBOX")
                    return EnumControlType.COMBOBOX;
                else if (_update.ControlType == "GRID_COMBOBOX")
                    return EnumControlType.GRID_COMBOBOX;
                else if (_update.ControlType == "GRID_TEXTBOX")
                    return EnumControlType.GRID_TEXTBOX;
                else if (_update.ControlType == "GRID_TEXTBOXDROPDOWN")
                    return EnumControlType.GRID_TEXTBOXDROPDOWN;
                else if (_update.ControlType == "RADIO_BUTTON")
                    return EnumControlType.RADIO_BUTTON;
                else if (_update.ControlType == "TEXTBOX")
                    return EnumControlType.TEXTBOX;
                else
                    return EnumControlType.TEXTBOX;
            }

            set 
            {
                _update.ControlType = value.ToString();
            }
        }


        /// <summary>
        /// 題目類型
        /// </summary>
        public EnumQuestionType QQuestionType 
        {
            get 
            {
                if (_update.QuestionType == "MULTI_ANSWER")
                    return EnumQuestionType.MULTI_ANSWER;
                else if (_update.QuestionType == "SEMESTER")
                    return EnumQuestionType.SEMESTER;
                else if (_update.QuestionType == "SINGLE_ANSWER")
                    return EnumQuestionType.SINGLE_ANSWER;
                else if (_update.QuestionType == "YEARLY")
                    return EnumQuestionType.YEARLY;
                else
                    return EnumQuestionType.SINGLE_ANSWER;

            }
            set 
            {
                _update.QuestionType = value.ToString();
            }
        }

        /// <summary>
        /// 取得更新後資料
        /// </summary>
        /// <returns></returns>
        public UDTQuestionsDataDef GetUpdateData()
        {
            return _update;
        }

        /// <summary>
        /// 是否能列印
        /// </summary>
        public bool CanPrint
        {
            get { return _update.CanPrint; }
            set { _update.CanPrint = value; }
        }

        public bool CanStudentEdit
        {
            get { return _update.CanStudentEdit; }
            set { _update.CanStudentEdit = value; }        
        }

        public bool CanTeacherEdit
        {
            get { return _update.CanTeacherEdit; }
            set { _update.CanTeacherEdit = value; }
        }

        public string ControlType
        {
            get { return _update.ControlType; }
            set { _update.ControlType = value; }
        }

        public int? DisplayOrder
        {
            get { return _update.DisplayOrder; }
            set { _update.DisplayOrder = value; }
        }

        public string Items
        {
            get { return _update.Items; }
            set { _update.Items = value; }
        }

        public string QuestionType
        {
            get { return _update.QuestionType; }
            set { _update.QuestionType = value; }        
        }

        public List<QuestionItem> itemList         
        {
            get 
            {
                List<QuestionItem> retVal = new List<QuestionItem>();
                if (_update.Items != "")
                {
                    XElement root = XElement.Parse(_update.Items);
                    if (root != null)
                    {
                        foreach (XElement elm in root.Elements("item"))
                        {
                            QuestionItem qi = new QuestionItem(elm.Attribute("key").Value,bool.Parse(elm.Attribute("has_remark").Value));
                            retVal.Add(qi);
                        }
                    }
                }
                return retVal;
            }

            set
            {
                if (value.Count > 0)
                {
                    XElement root = new XElement("Items");
                    foreach (QuestionItem item in value)
                    {
                        XElement elm = new XElement("item");
                        elm.SetAttributeValue("key", item.Key);
                        elm.SetAttributeValue("has_remark", item.hasRemark.ToString());
                        root.Add(elm);
                    }
                    _update.Items = root.ToString();
                }
            }
        
        }


    }
}
