using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// 輔導-親屬資訊
    /// </summary>
    [TableName("ischool.counsel.relative")]
    public class UDTRelativeDef:ActiveRecord
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
        /// 存歿
        ///</summary>
        [Field(Field = "is_alive", Indexed = false)]
        public bool? IsAlive { get; set; }

        ///<summary>
        /// 出生年
        ///</summary>
        [Field(Field = "birth_year", Indexed = false)]
        public int? BirthYear { get; set; }

        ///<summary>
        /// 職業
        ///</summary>
        [Field(Field = "job", Indexed = false)]
        public string Job { get; set; }

        ///<summary>
        /// 工作機構
        ///</summary>
        [Field(Field = "institute", Indexed = false)]
        public string Institute { get; set; }

        ///<summary>
        /// 職稱
        ///</summary>
        [Field(Field = "job_title", Indexed = false)]
        public string JobTitle { get; set; }

        ///<summary>
        /// 教育程度
        ///</summary>
        [Field(Field = "edu_degree", Indexed = false)]
        public string EduDegree { get; set; }

        /// <summary>
        /// 電話
        /// </summary>
        [Field(Field = "phone", Indexed = false)]
        public string Phone { get; set; }

        /// <summary>
        /// 原國籍
        /// </summary>
        [Field(Field = "national", Indexed = false)]
        public string National { get; set; }

        /// <summary>
        /// 行動電話
        /// </summary>
        [Field(Field = "cell_phone", Indexed = false)]
        public string CellPhone { get; set; }
    }
}
