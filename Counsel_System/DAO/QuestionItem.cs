using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Counsel_System2.DAO
{
    public class QuestionItem
    {
        public string Key { get; set; }

        public bool hasRemark { get; set; }

        public QuestionItem(string key, bool HasRemak)
        {
            Key=key;
            hasRemark=HasRemak;
        }

        public QuestionItem(string key)
            : this(key, false)
        {
            Key = key;
        }
    }
}
