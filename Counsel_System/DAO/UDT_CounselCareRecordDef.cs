using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// 輔導-優先關懷
    /// </summary>
    [TableName("counsel.care_record")]
    public class UDT_CounselCareRecordDef:ActiveRecord
    {
        /// <summary>
        ///  學生編號
        /// </summary>
        [Field(Field = "ref_student_id", Indexed = false)]
        public int StudentID { get; set; }

        /// <summary>
        ///  代號
        /// </summary>
        [Field(Field = "code_name", Indexed = false)]
        public string CodeName { get; set; }

        /// <summary>
        ///  立案日期
        /// </summary>
        [Field(Field = "file_date", Indexed = false)]
        public DateTime? FileDate { get; set; }

        /// <summary>
        ///  個案類別
        /// </summary>
        [Field(Field = "case_category", Indexed = false)]
        public string CaseCategory { get; set; }


        /// <summary>
        ///  個案類別備註
        /// </summary>
        [Field(Field = "case_category_remark", Indexed = false)]
        public string CaseCategoryRemark { get; set; }


        /// <summary>
        ///  個案來源
        /// </summary>
        [Field(Field = "case_origin", Indexed = false)]
        public string CaseOrigin { get; set; }

        /// <summary>
        ///  個案來源備註
        /// </summary>
        [Field(Field = "case_origin_remark", Indexed = false)]
        public string CaseOriginRemark { get; set; }


        /// <summary>
        ///  優勢
        /// </summary>
        [Field(Field = "superiority", Indexed = false)]
        public string Superiority { get; set; }

        /// <summary>
        ///  弱勢
        /// </summary>
        [Field(Field = "weakness", Indexed = false)]
        public string Weakness { get; set; }

        /// <summary>
        ///  校外輔導機構
        /// </summary>
        [Field(Field = "other_institute", Indexed = false)]
        public string OtherInstitute { get; set; }


        /// <summary>
        ///  輔導目標
        /// </summary>
        [Field(Field = "counsel_goal", Indexed = false)]
        public string CounselGoal { get; set; }

        /// <summary>
        ///  輔導方式
        /// </summary>
        [Field(Field = "counsel_type", Indexed = false)]
        public string CounselType { get; set; }

        /// <summary>
        ///  協助導師事項
        /// </summary>
        [Field(Field = "assisted_matter", Indexed = false)]
        public string AssistedMatter { get; set; }

        /// <summary>
        /// 記錄人的登入帳號
        /// </summary>
        [Field(Field = "author_id", Indexed = false)]
        public string AuthorID { get; set; }

        /// <summary>
        /// 記錄人的姓名
        /// </summary>
        [Field(Field = "author_name", Indexed = false)]
        public string AuthorName { get; set; }


        /// <summary>
        /// 學生狀態(不存入UDT)
        /// </summary>
        public string StudentStatus { get; set; }
    }
}
