using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// 測驗資料
    /// </summary>
    [TableName("counsel.quiz")]
    public class UDT_QuizDef : ActiveRecord
    {

        /// <summary>
        /// 測驗名稱
        /// </summary>        
        [Field(Field = "quiz_name", Indexed = false)]
        public string QuizName { get; set; }

        /// <summary>
        /// 測驗資料欄位(XML)
        /// </summary>        
        [Field(Field = "quiz_data_field", Indexed = false)]
        public string QuizDataField { get; set; }

        /// <summary>
        /// 記錄人的登入帳號
        /// </summary>
        [Field(Field = "author_id", Indexed = false)]
        public string AuthorID { get; set; }
    }
}
