using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace Counsel_System.DAO
{
    // <summary>
    /// 輔導-綜合記錄表-曾患特殊疾病
    /// </summary>
    [TableName("ischool.counsel.special_disease")]
    public class UDT_AB_SpecialDisease:ActiveRecord
    {
        ///<summary>
        /// 學生系統編號
        ///</summary>
        [Field(Field = "ref_student_id", Indexed = false)]
        public int StudentID { get; set; }

        ///<summary>
        /// 疾病名稱
        ///</summary>
        [Field(Field = "disease", Indexed = true)]
        public string Disease { get; set; }

    }
}
