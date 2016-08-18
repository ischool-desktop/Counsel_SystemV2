using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// 輔導-複選記錄
    /// </summary>
    [TableName("ischool.counsel.multiple_record")]
    public class UDTMultipleRecordDef:ActiveRecord
    {
        ///<summary>
        /// 學生編號
        ///</summary>
        [Field(Field = "ref_student_id", Indexed = false)]
        public int StudentID { get; set; }

        ///<summary>
        /// 識別值
        ///</summary>
        [Field(Field = "key", Indexed = false)]
        public string Key { get; set; }

        ///<summary>
        /// 資料
        ///</summary>
        [Field(Field = "data", Indexed = false)]
        public string Data { get; set; }

        ///<summary>
        /// 備註
        ///</summary>
        [Field(Field = "remark", Indexed = false)]
        public string Remark { get; set; }


    }
}
