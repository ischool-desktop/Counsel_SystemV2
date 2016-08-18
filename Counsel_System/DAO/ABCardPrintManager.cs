using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// 綜合表現紀錄表管理者
    /// </summary>
    public class ABCardPrintManager
    {
        Dictionary<int, UDT_ABCardTemplateDefinitionDef> _ABCardTemplateDict;
        Dictionary<string, string> _DataDict;
        Dictionary<string, List<AnswerPkey>> _AnswerDict;
        List<string> _MappingField;

        public ABCardPrintManager()
        {
            _DataDict = new Dictionary<string, string>();
            _MappingField = new List<string>();
            _ABCardTemplateDict = new Dictionary<int, UDT_ABCardTemplateDefinitionDef>();
            _AnswerDict = new Dictionary<string, List<AnswerPkey>>();

            foreach (UDT_ABCardTemplateDefinitionDef data in UDTTransfer.GetABCardTemplate())
            {
                _ABCardTemplateDict.Add(int.Parse(data.UID), data);
            }
        }

        /// <summary>
        /// 取得題目合併欄位
        /// </summary>
        /// <returns></returns>
        public List<string> GetMappingField()
        {
            return _MappingField;
        }

        /// <summary>
        /// 取得答案
        /// </summary>
        /// <returns></returns>
        public Dictionary<string, List<AnswerPkey>> GetAnswers(List<string> StudentIDList)
        {
            _MappingField.Clear();
            GetStudentABCardAnswerDict(StudentIDList);

            // 取得題目
            List<DAO.QuestionPKey> Questions = Global._ABCardTemplateTransfer.GetAllQuestionList();

            try
            {
                foreach (QuestionPKey question in Questions)
                {
                    if (question.QType == "grid")
                    {
                        //int RowCount = question.dataElement.Elements("Item").Count();
                        //// 當 grid 再細分
                        //foreach (XElement field in question.dataElement.Element("Cols").Elements("Col"))
                        //{
                        //    string key = question.GetField() + "_" + field.Attribute("name").Value + "R";
                        //    if (!_MappingField.Contains(key))
                        //        _MappingField.Add(key);
                        //}
                    }
                    else
                    {
                        if (!_MappingField.Contains(question.GetFieldG()))
                            _MappingField.Add(question.GetFieldG());

                        if (!_MappingField.Contains(question.GetField()))
                            _MappingField.Add(question.GetField());
                    }
                
                }

                

                // 題目與答案 mapping
                foreach (List<AnswerPkey> apkList in _AnswerDict.Values)
                {
                    foreach (AnswerPkey apk in apkList)
                    {
                        foreach (DAO.QuestionPKey question in Questions)
                        {
                            if (apk.TemplateID == question.TemplateID && apk.Label == question.SubjectLabel && apk.Name == question.Q_Name)
                            {
                                apk.Question = question;

                                //if (question.QType == "grid")
                                //{
                                //    int RowCount = apk.dataElement.Elements("Item").Count();
                                //    // 當 grid 再細分
                                //    foreach (XElement field in question.dataElement.Element("Cols").Elements("Col"))
                                //    {
                                //        string key = question.GetField() + "_" + field.Attribute("name").Value + "R";
                                //        if (!_MappingField.Contains(key))
                                //            _MappingField.Add(key);
                                //    }
                                //}
                                //else
                                //{
                                //    if(!_MappingField.Contains(question.GetFieldG()))
                                //        _MappingField.Add(question.GetFieldG());

                                //    if (!_MappingField.Contains(question.GetField()))
                                //        _MappingField.Add(question.GetField());
                                //}


                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                FISCA.Presentation.Controls.MsgBox.Show("答案解析失敗." + ex.Message);
            }

            return _AnswerDict;
        }

        /// <summary>
        /// 取得學生綜合表現答案
        /// </summary>
        /// <param name="StudentIDList"></param>
        private void GetStudentABCardAnswerDict(List<string> StudentIDList)
        {
            _AnswerDict.Clear();
            // 取得學生綜合表現答案
            List<UDT_ABCardDataDef> ABCardDataList = UDTTransfer.GetABCardDataListByStudentList(StudentIDList);
            
            foreach (string str in StudentIDList)
            { 
                int id=int.Parse(str);
                List<AnswerPkey> dataList = new List<AnswerPkey>();
                foreach (UDT_ABCardDataDef data in ABCardDataList.Where(x => x.StudentID == id))
                {
                    XElement ansElmRoot = XElement.Parse(data.Content);

                    foreach (XElement elm in ansElmRoot.Elements("Ans"))
                    {
                        AnswerPkey apk = new AnswerPkey();
                        apk.dataElement = elm;
                        apk.Label = data.SubjectName;
                        apk.Name = elm.Attribute("name").Value;
                        apk.StudentID = id;
                        apk.TemplateID = data.TemplateID;
                        dataList.Add(apk);
                    }
                }

                if (dataList.Count > 0)
                {
                    _AnswerDict.Add(str, dataList);                
                }
            }
        }

    }
}
