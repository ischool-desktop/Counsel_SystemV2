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
    [TableName("counsel.interview_record")]
    public class UDT_CounselStudentInterviewRecordDef:ActiveRecord
    {

        /// <summary>
        ///  晤談編號(使用者自行輸入)
        /// </summary>
        [Field(Field = "interview_no", Indexed = false)]
        public string InterviewNo { get; set; }

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
        /// 晤談方式
        /// </summary>
        [Field(Field = "interview_type", Indexed = false)]
        public string InterviewType { get; set; }

        /// <summary>
        /// 晤談對象
        /// </summary>
        [Field(Field = "interviewee_type", Indexed = false)]
        public string IntervieweeType { get; set; }

        /// <summary>
        /// 晤談日期
        /// </summary>
        [Field(Field = "interview_date", Indexed = false)]
        public DateTime? InterviewDate { get; set; }


        /// <summary>
        /// 晤談時間
        /// </summary>
        [Field(Field = "interview_time", Indexed = false)]
        public string InterviewTime { get; set; }


        /// <summary>
        /// 地點
        /// </summary>
        [Field(Field = "place", Indexed = false)]
        public string Place { get; set; }

        /// <summary>
        /// 事由
        /// </summary>
        [Field(Field = "cause", Indexed = false)]
        public string Cause { get; set; }

        /// <summary>
        /// 存放參與人員(XML)
        /// </summary>
        [Field(Field = "attendees", Indexed = false)]
        public string Attendees { get; set; }

        /// <summary>
        /// 存放輔導方式(XML)
        /// </summary>
        [Field(Field = "counsel_type", Indexed = false)]
        public string CounselType { get; set; }

        /// <summary>
        /// 存放輔導歸類(XML)
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


        /// <summary>
        /// 學生狀態(不存入UDT)
        /// </summary>
        public string StudentStatus { get; set; }

    }
}
