using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace Counsel_System2.DAO
{

    /// <summary>
    /// 晤談紀錄
    /// </summary>
    [TableName("counsel.home_visit_record")]
    public class UDT_Counsel_home_visit_RecordDef : ActiveRecord
    {
        /// <summary>
        ///  聯繫編號(使用者自行輸入)
        /// </summary>
        [Field(Field = "home_visit_no", Indexed = false)]
        public string home_visit_no { get; set; }

        /// <summary>
        ///  學年度
        /// </summary>
        [Field(Field = "schoolyear", Indexed = false)]
        public string schoolyear { get; set; }

        /// <summary>
        ///  學期
        /// </summary>
        [Field(Field = "semester", Indexed = false)]
        public string semester { get; set; }



        /// <summary>
        ///  記錄人
        /// </summary>
        [Field(Field = "author_role", Indexed = false)]
        public string authorRole { get; set; }



        /// <summary>
        /// 學生編號
        /// </summary>
        [Field(Field = "ref_student_id", Indexed = false)]
        public int StudentID { get; set; }

        /// <summary>
        /// 是否公開
        /// </summary>
        [Field(Field = "is_public", Indexed = false)]
        public bool isPublic { get; set; }

        /// <summary>
        /// 晤談老師編號(與TeacherID關聯)
        /// </summary>
        [Field(Field = "teacher_id", Indexed = false)]
        public int TeacherID { get; set; }

        /// <summary>
        /// 聯繫方式 
        /// </summary>
        [Field(Field = "home_visit_type", Indexed = false)]
        public string home_visit_type { get; set; }



        /// <summary>
        /// 聯繫日期
        /// </summary>
        [Field(Field = "home_visit_date", Indexed = false)]
        public DateTime? home_visit_date { get; set; }


        /// <summary>
        /// 聯繫時間
        /// </summary>
        [Field(Field = "home_visit_time", Indexed = false)]
        public string home_visit_time { get; set; }


        /// <summary>
        /// 地點
        /// </summary>
        [Field(Field = "place", Indexed = false)]
        public string Place { get; set; }

        /// <summary>
        /// 聯繫事由
        /// </summary>
        [Field(Field = "cause", Indexed = false)]
        public string Cause { get; set; }

        /// <summary>
        /// 參與成員(XML)
        /// </summary>
        [Field(Field = "attendees", Indexed = false)]
        public string Attendees { get; set; }

        /// <summary>
        /// 聯繫成員(XML)
        /// </summary>
        [Field(Field = "contact", Indexed = false)]
        public string contact { get; set; }

        /// <summary>
        /// 聯繫類別(XML)
        /// </summary>
        [Field(Field = "counsel_type_kind", Indexed = false)]
        public string CounselTypeKind { get; set; }


        /// <summary>
        /// 內容要點
        /// </summary>
        [Field(Field = "content_digest", Indexed = false)]
        public string ContentDigest { get; set; }

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


        ///// <summary>
        ///// 學生狀態(不存入UDT)
        ///// </summary>
        //public string StudentStatus { get; set; }


    }
}
