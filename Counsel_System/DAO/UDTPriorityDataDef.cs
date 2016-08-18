using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// 輔導-優先順序資料
    /// </summary>
     [TableName("ischool.counsel.priority_data")]
    public class UDTPriorityDataDef:ActiveRecord
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
        /// 順序1
        ///</summary>
        [Field(Field = "p1", Indexed = false)]
        public string P1 { get; set; }

        ///<summary>
        /// 順序2
        ///</summary>
        [Field(Field = "p2", Indexed = false)]
        public string P2 { get; set; }

        ///<summary>
        /// 順序3
        ///</summary>
        [Field(Field = "p3", Indexed = false)]
        public string P3 { get; set; }

        ///<summary>
        /// 順序4
        ///</summary>
        [Field(Field = "p4", Indexed = false)]
        public string P4 { get; set; }

        ///<summary>
        /// 順序5
        ///</summary>
        [Field(Field = "p5", Indexed = false)]
        public string P5 { get; set; }

        ///<summary>
        /// 順序6
        ///</summary>
        [Field(Field = "p6", Indexed = false)]
        public string P6 { get; set; }

        ///<summary>
        /// 順序7
        ///</summary>
        [Field(Field = "p7", Indexed = false)]
        public string P7 { get; set; }

        ///<summary>
        /// 順序8
        ///</summary>
        [Field(Field = "p8", Indexed = false)]
        public string P8 { get; set; }

        ///<summary>
        /// 順序9
        ///</summary>
        [Field(Field = "p9", Indexed = false)]
        public string P9 { get; set; }

        ///<summary>
        /// 順序10
        ///</summary>
        [Field(Field = "p10", Indexed = false)]
        public string P10 { get; set; }

    }
}
