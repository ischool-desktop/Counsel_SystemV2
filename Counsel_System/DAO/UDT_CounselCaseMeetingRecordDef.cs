using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// 輔導-個案會議
    /// </summary>
    [TableName("counsel.case_meeting_record")]
    public class UDT_CounselCaseMeetingRecordDef:ActiveRecord
    {
            /// <summary>
            ///  學生編號
            /// </summary>
            [Field(Field = "ref_student_id", Indexed = false)]
            public int StudentID { get; set; }


            /// <summary>
            ///  個案編號
            /// </summary>
            [Field(Field = "case_no", Indexed = false)]
            public string CaseNo { get; set; }


            /// <summary>
            ///  晤談老師
            /// </summary>
            [Field(Field = "counsel_teacher_id", Indexed = false)]
            public int CounselTeacherID { get; set; }


            /// <summary>
            ///  會議日期
            /// </summary>
            [Field(Field = "meeting_date", Indexed = false)]
            public DateTime? MeetingDate { get; set; }

            /// <summary>
            ///  會議時間
            /// </summary>
            [Field(Field = "meeting_time", Indexed = false)]
            public string MeetigTime { get; set; }


            /// <summary>
            ///  晤談地點
            /// </summary>
            [Field(Field = "place", Indexed = false)]
            public string Place { get; set; }


            /// <summary>
            ///  參與人員 , xml
            /// </summary>
            [Field(Field = "attendees", Indexed = false)]
            public string Attendees { get; set; }

            /// <summary>
            ///  會議事由
            /// </summary>
            [Field(Field = "meeting_cause", Indexed = false)]
            public string MeetingCause { get; set; }

            /// <summary>
            ///  輔導方式 , xml
            /// </summary>
            [Field(Field = "counsel_type", Indexed = false)]
            public string CounselType { get; set; }

            /// <summary>
            ///  輔導歸類 , xml
            /// </summary>
            [Field(Field = "counsel_type_kind", Indexed = false)]
            public string CounselTypeKind { get; set; }

            /// <summary>
            ///  內容要點
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
