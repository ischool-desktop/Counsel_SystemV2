using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace Counsel_System2.Parser
{
    public class Subject
    {
        private string label = "";
        private Dictionary<string, QuestionGroup> questionGroups;
        private List<QuestionGroup> qgs;

        public Subject(XmlElement elmSubject)
        {
            this.questionGroups = new Dictionary<string,QuestionGroup>();
            this.qgs = new List<QuestionGroup>();

            foreach (XmlElement elm in elmSubject.SelectNodes("QG"))
            {
                QuestionGroup qg = new QuestionGroup(elm);
                this.questionGroups.Add(qg.GetIdentifyLabel(), qg);
                this.qgs.Add(qg);
            }
        }

        public string GetLabel()
        {
            return this.label;
        }

        public List<QuestionGroup> GetQuestionGroups()
        {
            return this.qgs;
        }

        public Dictionary<string, QuestionGroup> GetDicQuestionGroups()
        {
            return this.questionGroups;
        }
    }
}
