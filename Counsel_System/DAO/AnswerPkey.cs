using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// 綜合表現答案
    /// </summary>
    public class AnswerPkey
    {
        /// <summary>
        /// 學生系統編號
        /// </summary>
        public int StudentID { get; set; }

        /// <summary>
        /// 答案標籤
        /// </summary>
        public string Label { get; set; }

        /// <summary>
        /// 樣板ID
        /// </summary>
        public int TemplateID { get; set; }

        /// <summary>
        /// 答案KeyName
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 答案資料
        /// </summary>
        public XElement dataElement { get; set; }

        /// <summary>
        /// 放題目
        /// </summary>
        public DAO.QuestionPKey Question { get; set; }
    }
}
