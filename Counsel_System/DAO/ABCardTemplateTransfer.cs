using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// 處理綜合表現樣板交換
    /// </summary>
    public class ABCardTemplateTransfer
    {
        private List<string> _SubjectQuestionNames;

        List<DAO.UDT_ABCardTemplateDefinitionDef> _ABCardTemplate;
        Dictionary<string, XElement> _SubjectDict;
        Dictionary<string, string> _QTypeNameDict;
        Dictionary<string, string> _QTypeKeyDict;
        Dictionary<string, int> _TemplateIDDict;

        public ABCardTemplateTransfer()
        {
            _QTypeKeyDict = new Dictionary<string, string>();
            _QTypeNameDict = new Dictionary<string, string>();

            _TemplateIDDict = new Dictionary<string, int>();
        }

        /// <summary>
        /// 取得問題類型
        /// </summary>
        /// <returns></returns>
        public static List<string> GetQTypeList()
        {
            List<string> _QTypeList= new List<string>();
            _QTypeList.Add("grid");
            _QTypeList.Add("combobox");
            _QTypeList.Add("text");
            _QTypeList.Add("textboxdropdown");
            _QTypeList.Add("checkbox");
            _QTypeList.Add("textarea");

            return _QTypeList;        
        }

        /// <summary>
        /// 載入所有樣板
        /// </summary>
        /// <param name="ABCardTemplate"></param>
        public void LoadAllTemplate(List<DAO.UDT_ABCardTemplateDefinitionDef> ABCardTemplate)
        {
            _SubjectQuestionNames = new List<string>();

            // 取得樣板
            _ABCardTemplate = ABCardTemplate;

            // 全部題目
            _SubjectDict = new Dictionary<string, XElement>();

            // 解析 XML
            foreach (DAO.UDT_ABCardTemplateDefinitionDef data in _ABCardTemplate)
            {
                XElement elm = XElement.Parse(data.Content);
                _SubjectDict.Add(data.SubjectName, elm);

                if (!_TemplateIDDict.ContainsKey(data.SubjectName))
                    _TemplateIDDict.Add(data.SubjectName, int.Parse(data.UID));
            }        
        }

        /// <summary>
        /// 透過 SubjectName 取得所有 QGLabelName
        /// </summary>
        /// <param name="SubjectName"></param>
        /// <returns></returns>
        public List<string> GetQGLabelList(string SubjectLabel)
        {
            List<string> retVal = new List<string>();

            if (_SubjectDict.ContainsKey(SubjectLabel))
                retVal = (from data in _SubjectDict[SubjectLabel].Elements("QG") select data.Attribute("label").Value).ToList();

            return retVal;
        }

        /// <summary>
        /// 透過 SubjectLabel 取得 UDT 內資料(Cache)
        /// </summary>
        /// <param name="SubjectLabel"></param>
        /// <returns></returns>
        public DAO.UDT_ABCardTemplateDefinitionDef GetUDTTemplateBySubjectLabel(string SubjectLabel)
        {
            DAO.UDT_ABCardTemplateDefinitionDef retVal = null;

            foreach (DAO.UDT_ABCardTemplateDefinitionDef data in _ABCardTemplate)
            {
                if (data.SubjectName == SubjectLabel)
                    retVal = data;
            }
            return retVal;
        }


        /// <summary>
        /// 透過 SubjectLabel 取得題目
        /// </summary>
        /// <param name="SubjectLabel"></param>
        /// <returns></returns>
        public List<QuestionPKey> GetQuestionListBySubjectLabel(string SubjectLabel)
        {
            List<QuestionPKey> retVal = new List<QuestionPKey>();
            if (_SubjectDict.ContainsKey(SubjectLabel))
            {
                foreach (XElement elm in _SubjectDict[SubjectLabel].Elements("QG"))
                {
                    if (elm.Attribute("label") != null)
                        foreach (XElement elmQ in elm.Element("Qs").Elements("Q"))
                        {
                            QuestionPKey QPK = new QuestionPKey();
                            QPK.dataElement = elmQ;
                            QPK.Q_Name = elmQ.Attribute("name").Value;
                            QPK.QGLabel = elm.Attribute("label").Value;
                            QPK.QLabel = elmQ.Attribute("label").Value;
                            QPK.QType = elmQ.Attribute("type").Value;
                            QPK.TemplateID = _TemplateIDDict[SubjectLabel];
                            QPK.SubjectLabel = SubjectLabel;
                            retVal.Add(QPK);
                        }
                }
            }
            return retVal;
        }

        /// <summary>
        /// 取得所有題目
        /// </summary>
        /// <returns></returns>
        public List<QuestionPKey> GetAllQuestionList()
        {
            List<string> ChoiceTmp = new List<string> ();
            List<QuestionPKey> retVal = new List<QuestionPKey>();
            foreach(string SubjectLabel in _SubjectDict.Keys)
            {
                foreach (XElement elm in _SubjectDict[SubjectLabel].Elements("QG"))
                {
                    ChoiceTmp.Clear();
                    if(elm.Element("Choices")!=null)
                    {
                        foreach(XElement elmCho in elm.Element("Choices").Elements("Item"))
                            ChoiceTmp.Add(elmCho.Attribute("label").Value);
                    }

                    if (elm.Attribute("label") != null)
                        foreach (XElement elmQ in elm.Element("Qs").Elements("Q"))
                        {
                            QuestionPKey QPK = new QuestionPKey();
                            QPK.dataElement = elmQ;
                            QPK.Q_Name = elmQ.Attribute("name").Value;
                            QPK.QGLabel = elm.Attribute("label").Value;
                            QPK.QLabel = elmQ.Attribute("label").Value;
                            QPK.QType = elmQ.Attribute("type").Value;
                            QPK.TemplateID = _TemplateIDDict[SubjectLabel];
                            QPK.SubjectLabel = SubjectLabel;
                            QPK.Q_Choice = "";
                            if (ChoiceTmp.Count > 0)
                                QPK.Q_Choice = string.Join(";", ChoiceTmp.ToArray());

                            retVal.Add(QPK);
                        }
                }
            }
            return retVal;
        }




        /// <summary>
        /// 透過SubjectName 取得所有 QsQName
        /// </summary>
        /// <param name="SubjectName"></param>
        /// <returns></returns>
        public List<string> GetQsQNameList(string SubjectLabel)
        {
            List<string> retVal = new List<string>();
            if (_SubjectDict.ContainsKey(SubjectLabel))
            {
                foreach (XElement elmQG in _SubjectDict[SubjectLabel].Elements("QG"))
                {
                    foreach (XElement Qsq in elmQG.Element("Qs").Elements("Q"))
                        retVal.Add(Qsq.Attribute("name").Value);                
                }
            }
            return retVal;
        }

        


        /// <summary>
        /// 取得所有Question Name 給比對用
        /// </summary>
        /// <returns></returns>
        public List<string> GetAllQsQNameList()
        {
            List<string> retVal = new List<string>();
            foreach(string str in _SubjectDict.Keys)
            {
                foreach (XElement elmQG in _SubjectDict[str].Elements("QG"))
                {
                    foreach (XElement Qsq in elmQG.Element("Qs").Elements("Q"))
                        retVal.Add(Qsq.Attribute("name").Value);
                }
            }
            return retVal;        
        }


        /// <summary>
        /// 取得所選擇
        /// </summary>
        /// <param name="SubjectName"></param>
        /// <param name="QGLabel"></param>
        /// <returns></returns>
        public XElement GetChoicesByQGLabel(string SubjectLabel,string QGLabel)
        {
            XElement retVal = null;
            if (_SubjectDict.ContainsKey(SubjectLabel))
                retVal = _SubjectDict[SubjectLabel].Element(QGLabel).Element("Choice");          

            return retVal;
        }

        /// <summary>
        /// 透過特定Key組合法取得單一Question
        /// </summary>
        /// <param name="qPkey"></param>
        /// <param name="QType"></param>
        /// <returns></returns>
        public XElement GetQsQByKey1(QuestionPKey qPkey)
        {
            XElement retVal = null;
            if (_SubjectDict.ContainsKey(qPkey.SubjectLabel))
            {
                foreach (XElement elm in _SubjectDict[qPkey.SubjectLabel].Elements("QG"))
                { 
                    if(elm.Attribute("label") !=null )
                        if (elm.Attribute("label").Value == qPkey.QGLabel)
                        {
                            foreach (XElement elmQ in elm.Element("Qs").Elements("Q"))
                            {
                                if (elmQ.Attribute("label").Value == qPkey.QLabel && elmQ.Attribute("type").Value == qPkey.QType)
                                {
                                    retVal = elmQ;                                
                                }                            
                            }                        
                        }               
                }            
            }
            return retVal;
        }


        /// <summary>
        /// 透過特定Key組合法存入Question
        /// </summary>
        /// <param name="qPkey"></param>
        /// <param name="QType"></param>
        /// <param name="data"></param>
        public void SetQsQElement(QuestionPKey qPkey)
        {
            if (_SubjectDict.ContainsKey(qPkey.SubjectLabel))
            {
                foreach (XElement elm in _SubjectDict[qPkey.SubjectLabel].Elements("QG"))
                {
                    if (elm.Attribute("label") != null)
                        if (elm.Attribute("label").Value == qPkey.QGLabel)
                        {
                            foreach (XElement elmQ in elm.Element("Qs").Elements("Q"))
                            {
                                if (elmQ.Attribute("label").Value == qPkey.QLabel && elmQ.Attribute("type").Value == qPkey.QType)
                                {
                                    elm.Remove();
                                }
                            }
                            elm.Element("Qs").Add(qPkey.dataElement);
                        }
                }
                
            }       
        
        }


        /// <summary>
        /// 是否有題組
        /// </summary>
        /// <param name="SubjectName"></param>
        /// <returns></returns>
        public bool HasSubject(string SubjectName)
        {
            bool retVal = false;
            if (_SubjectDict.ContainsKey(SubjectName))
                retVal = true;

            return retVal;
        }

        /// <summary>
        /// 取得所有題組名稱
        /// </summary>
        /// <returns></returns>
        public List<string> GetAllSubjectNameList()
        {
            List<string> retVal = new List<string>();
            retVal = _SubjectDict.Keys.ToList();
            retVal.Sort();
            return retVal;
        }
    }
}
