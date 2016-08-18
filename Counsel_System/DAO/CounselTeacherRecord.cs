using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// 輔導系統老師
    /// </summary>
    public class CounselTeacherRecord
    {
        public enum CounselTeacherType {輔導老師,認輔老師,輔導主任}

        /// <summary>
        /// 老師完整名稱TeacherName(NickName)
        /// </summary>
        public string TeacherFullName { get; set; }

        /// <summary>
        /// TeacherTagID(資料庫存取)
        /// </summary>
        public int TeacherTag_ID { get; set; }

        /// <summary>
        /// 教師編號
        /// </summary>
        public string TeacherID { get; set; }

        /// <summary>
        /// 輔導系統老師類別
        /// </summary>
        public CounselTeacherType counselTeacherType { get; set; }
    }
}
