using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace Counsel_System.DAO
{
    /// <summary>
    /// 輔導-綜合記錄表-兄弟姊妹資訊
    /// </summary>
    [TableName("ischool.counsel.sibling")]
    public class UDT_AB_Sibling:ActiveRecord
    {
        ///<summary>
        /// 學生編號
        ///</summary>
        [Field(Field = "ref_student_id", Indexed = false)]
        public int StudentID { get; set; }

        ///<summary>
        /// 稱謂
        ///</summary>
        [Field(Field = "title", Indexed = false)]
        public string Title { get; set; }

        ///<summary>
        /// 姓名
        ///</summary>
        [Field(Field = "name", Indexed = false)]
        public string Name { get; set; }

        ///<summary>
        /// 畢肆業學校
        ///</summary>
        [Field(Field = "school_name", Indexed = false)]
        public string SchoolName { get; set; }

        ///<summary>
        /// 出生年次
        ///</summary>
        [Field(Field = "birth_year", Indexed = false)]
        public int BirthYear { get; set; }

    }
}
