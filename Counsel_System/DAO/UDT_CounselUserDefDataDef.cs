using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace Counsel_System2.DAO
{
    // 輔導使用者自訂資料欄位
    [TableName("Counsel.UserDefineData")]
    public class UDT_CounselUserDefDataDef:ActiveRecord
    {
        /// <summary>
        /// 欄位名稱
        /// </summary>
        [Field(Field = "field_name", Indexed = false)]
        public string FieldName { get; set; }

        /// <summary>
        /// 資料型態
        /// </summary>
        [Field(Field = "data_type", Indexed = false)]
        public string DataType { get; set; }

        /// <summary>
        /// 資料值
        /// </summary>
        [Field(Field = "value", Indexed = false)]
        public string Value { get; set; }

        /// <summary>
        /// 學生ID
        /// </summary>
        [Field(Field = "ref_student_id", Indexed = false)]
        public int StudentID { get; set; }

        /// <summary>
        /// 學生狀態(不存入UDT)
        /// </summary>
        public string StudentStatus { get; set; }

    }
}
