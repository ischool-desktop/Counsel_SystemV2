using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// 輔導-每年資料
    /// </summary>
    [TableName("ischool.counsel.yearly_data")]
    public class UDTYearlyDataDef:ActiveRecord
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
        /// 1年級
        ///</summary>
        [Field(Field = "g1", Indexed = false)]
        public string G1 { get; set; }

        ///<summary>
        /// 2年級
        ///</summary>
        [Field(Field = "g2", Indexed = false)]
        public string G2 { get; set; }

        ///<summary>
        /// 3年級
        ///</summary>
        [Field(Field = "g3", Indexed = false)]
        public string G3 { get; set; }

        ///<summary>
        /// 4年級
        ///</summary>
        [Field(Field = "g4", Indexed = false)]
        public string G4 { get; set; }

        ///<summary>
        /// 5年級
        ///</summary>
        [Field(Field = "g5", Indexed = false)]
        public string G5 { get; set; }

        ///<summary>
        /// 6年級
        ///</summary>
        [Field(Field = "g6", Indexed = false)]
        public string G6 { get; set; }


    }
}
