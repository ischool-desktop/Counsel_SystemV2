using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// 綜合表現題目管理
    /// </summary>
    public class ABCardQuestionDataManager
    {
        List<QuestionData> _QuestionDataList;
        List<UDTQuestionsDataDef> _UDTQuestionsDataList;

        public ABCardQuestionDataManager()
        {
            _QuestionDataList = new List<QuestionData>();
            _UDTQuestionsDataList = UDTTransfer.ABUDTQuestionsDataSelectAll();
            //UDTTransfer.ABUDTQuestionsDataDelete(_UDTQuestionsDataList);

            //_UDTQuestionsDataList = UDTTransfer.ABUDTQuestionsDataSelectAll();

            foreach (UDTQuestionsDataDef UQ in _UDTQuestionsDataList)
            {
                QuestionData qd = new QuestionData(UQ);
                _QuestionDataList.Add(qd);
            }

            // 當沒有資料放入預設
            if (_UDTQuestionsDataList.Count < 1)
                AddDefaultQuestionItems();                
        }

        /// <summary>
        /// 取得所有題目
        /// </summary>
        /// <returns></returns>
        public Dictionary<string, List<QuestionData>> GetAllQuestionData()
        {
            Dictionary<string, List<QuestionData>> retVal = new Dictionary<string, List<QuestionData>>();
            foreach (QuestionData qd in _QuestionDataList)
            {
                string key = qd.Group + "_" + qd.Name;

                if (retVal.ContainsKey(key))
                    retVal[key].Add(qd);
                else
                {
                    List<QuestionData> data = new List<QuestionData>();
                    data.Add(qd);
                    retVal.Add(key, data);
                }            
            }
            return retVal;
        }

        /// <summary>
        /// 依群組名稱取的資料
        /// </summary>
        /// <param name="Name"></param>
        /// <returns></returns>
        public Dictionary<string, QuestionData> GetQuestionDataByGroupName(string Name)
        {
            Dictionary<string, QuestionData> retVal = new Dictionary<string, QuestionData>();
            List<UDTQuestionsDataDef> dataList = UDTTransfer.ABUDTQuestionsDataSelectByGroupName(Name);
            dataList=(from da in dataList orderby da.Group,da.Name select da).ToList();

            foreach (UDTQuestionsDataDef ud in dataList)
            {
                QuestionData qd = new QuestionData (ud);                
                if (!retVal.ContainsKey(qd.Name))
                    retVal.Add(qd.Name, qd);
            }

            return retVal;        
        }    
        
        private void ClearAll()
        {
            List<UDTQuestionsDataDef> delData = new List<UDTQuestionsDataDef>();
            foreach (QuestionData qd in _QuestionDataList)
                delData.Add(qd.GetUpdateData());
            UDTTransfer.ABUDTQuestionsDataDelete(delData);
        }

        private void AddDefaultQuestionItems()
        {
            List<QuestionData> qdList = new List<QuestionData>();
            
            XElement elmRoot = XElement.Parse(Properties.Resources.Questions);
            foreach (XElement elm in elmRoot.Elements("Question"))
            {
                QuestionData qd = new QuestionData(null);
                qd.Group = elm.Element("Group").Value;
                qd.Name = elm.Element("Name").Value;
                qd.QuestionType = elm.Element("QuestionType").Value;
                qd.ControlType = elm.Element("ControlType").Value;
                qd.CanPrint = bool.Parse(elm.Element("CanPrint").Value);
                qd.CanStudentEdit = bool.Parse(elm.Element("CanStudentEdit").Value);
                qd.CanTeacherEdit = bool.Parse(elm.Element("CanTeacherEdit").Value);
                qd.DisplayOrder = int.Parse(elm.Element("DisplayOrder").Value);
                if(elm.Element("Items") !=null)
                    qd.Items=elm.Element("Items").ToString();
                qdList.Add(qd);
            }
            InsertUDT(qdList);
        }

        private void InsertUDT(List<QuestionData> dataList)
        {
            List<UDTQuestionsDataDef> newData = new List<UDTQuestionsDataDef>();
            foreach (QuestionData qd in dataList)
                newData.Add(qd.GetUpdateData());

            UDTTransfer.ABUDTQuestionsDataInsert(newData);
        
        }

    }
}
