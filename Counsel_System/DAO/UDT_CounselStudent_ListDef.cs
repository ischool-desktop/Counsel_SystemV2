using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// 輔導學生與教師
    /// </summary>
    [TableName("counsel.student_list")]
    public class UDT_CounselStudent_ListDef:ActiveRecord
    {
        /// <summary>
        ///  教師清單編號，對應到 counsel.teacher_list.uid
        /// </summary>
        [Field(Field = "ref_teacher_tag_id", Indexed = false)]
        public int TeacherTagID { get; set; }

        /// <summary>
        ///  學生編號，對應到 student.id
        /// </summary>
        [Field(Field = "ref_student_id", Indexed = false)]
        public int StudentID { get; set; }

        /// <summary>
        /// 記錄人的登入帳號
        /// </summary>
        [Field(Field = "author_id", Indexed = false)]
        public string AuthorID { get; set; }

    }
}
