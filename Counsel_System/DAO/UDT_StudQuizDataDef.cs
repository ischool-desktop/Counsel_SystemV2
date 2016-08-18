using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// 學生測驗資料
    /// </summary>
    [TableName("counsel.student_quiz_data")]
    public class UDT_StudQuizDataDef : ActiveRecord
    {
        /// <summary>
        /// 測驗ID(Key)
        /// </summary>        
        [Field(Field = "ref_quiz_uid", Indexed = false)]
        public int QuizID { get; set; }

        /// <summary>
        /// 學生編號
        /// </summary>        
        [Field(Field = "ref_student_id", Indexed = false)]
        public int StudentID { get; set; }

        /// <summary>
        /// 實施日期
        /// </summary>        
        [Field(Field = "implementation_date", Indexed = false)]
        public DateTime? ImplementationDate { get; set; }

        /// <summary>
        /// 解析日期
        /// </summary>        
        [Field(Field = "analysis_date", Indexed = false)]
        public DateTime? AnalysisDate { get; set; }

        /// <summary>
        /// 內容(XML)
        /// </summary>        
        [Field(Field = "content", Indexed = false)]
        public string Content { get; set; }

        /// <summary>
        /// 記錄人的登入帳號
        /// </summary>
        [Field(Field = "author_id", Indexed = false)]
        public string AuthorID { get; set; }

        /// <summary>
        /// 學生狀態(不存入UDT)
        /// </summary>
        public string StudentStatus { get; set; }

    }
}
