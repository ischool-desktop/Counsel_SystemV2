using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace Counsel_System2.Parser
{
    /**
         * <Answers label="本人概況"> <!-- 對應到 Subject的 label -->
	            <Ans name="AAA12345670" value=”AB”></Ans>
	            <Ans name="AAA12345671" value=”基督教”></Ans>	            
	            <Ans name="AAA12345678" >
		            <Item value=”近視” />
                    <Item value=”其它” remark=”蛀牙” />
                </Ans>
	            <Ans name="AAA12345679">
                    <Item value=”氣喘” />
                    <Item value=”其它” remark=”身心症” />
                </Ans>
            </Answers>
         * */
    public class Answers
    {
        private string label = "";      //label of Subject
        private Dictionary<string, Answer> answers; //question name, Answer

        public Answers() {

        }

        public Answers(XmlElement elmAnswers)
        {
            this.label = elmAnswers.GetAttribute("label");
            this.answers = new Dictionary<string, Answer>();

            foreach (XmlElement elmAns in elmAnswers.SelectNodes("Ans"))
            {
                string qname = elmAns.GetAttribute("name");

                // 加入這段要處理當初版本 name bug


                if (!answers.ContainsKey(qname))
                    this.answers.Add(qname, new Answer(elmAns));
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show(qname);
                }
            }
        }


        public string GetSubjectLabel()
        {
            return this.label;
        }

        public void SetSubjectLabel(string subjLabel)
        {
            this.label = subjLabel;
        }

        /// <summary>
        /// 取得某一題號的答案
        /// </summary>
        /// <param name="qName"></param>
        /// <returns></returns>
        public Answer GetAnswer(string qName)
        {
            Answer result = null;
            if (this.answers.ContainsKey(qName))
                result = this.answers[qName];

            return result;
        }

        /// <summary>
        /// 取得此 Subject 的所有題目的答案
        /// </summary>
        /// <returns></returns>
        public List<Answer> GetAllAnswers()
        {
            return this.answers.Values.ToList<Answer>();
        }

    }
}
