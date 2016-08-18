using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// 輔導-每學期資料
    /// </summary>
    [TableName("ischool.counsel.semester_data")]
    public class UDTSemesterDataDef:ActiveRecord
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
        /// 1年級上學期
        ///</summary>
        [Field(Field = "s1a", Indexed = false)]
        public string S1a { get; set; }

        ///<summary>
        /// 1年級下學期
        ///</summary>
        [Field(Field = "s1b", Indexed = false)]
        public string S1b { get; set; }

        ///<summary>
        /// 2年級上學期
        ///</summary>
        [Field(Field = "s2a", Indexed = false)]
        public string S2a { get; set; }

        ///<summary>
        /// 2年級下學期
        ///</summary>
        [Field(Field = "s2b", Indexed = false)]
        public string S2b { get; set; }

        ///<summary>
        /// 3年級上學期
        ///</summary>
        [Field(Field = "s3a", Indexed = false)]
        public string S3a { get; set; }

        ///<summary>
        /// 3年級下學期
        ///</summary>
        [Field(Field = "s3b", Indexed = false)]
        public string S3b { get; set; }

        ///<summary>
        /// 4年級上學期
        ///</summary>
        [Field(Field = "s4a", Indexed = false)]
        public string S4a { get; set; }

        ///<summary>
        /// 4年級下學期
        ///</summary>
        [Field(Field = "s4b", Indexed = false)]
        public string S4b { get; set; }

        ///<summary>
        /// 5年級上學期
        ///</summary>
        [Field(Field = "s5a", Indexed = false)]
        public string S5a { get; set; }

        ///<summary>
        /// 5年級下學期
        ///</summary>
        [Field(Field = "s5b", Indexed = false)]
        public string S5b { get; set; }

        ///<summary>
        /// 6年級上學期
        ///</summary>
        [Field(Field = "s6a", Indexed = false)]
        public string S6a { get; set; }

        ///<summary>
        /// 6年級下學期
        ///</summary>
        [Field(Field = "s6b", Indexed = false)]
        public string S6b { get; set; }


    }
}
