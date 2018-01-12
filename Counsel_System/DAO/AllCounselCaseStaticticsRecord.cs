using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// 輔導系統老師
    /// </summary>
    public class AllCounselStaticticsTeacherRecord
    {
        public enum CounselTeacherType {班導師,輔導老師,認輔老師,輔導主任}

        /// <summary>
        /// 老師完整名稱TeacherName(NickName)
        /// </summary>
        public string TeacherName { get; set; }

        /// <summary>
        /// TeacherTagID(資料庫存取)
        /// </summary>
        public int TeacherTag_ID { get; set; }

        /// <summary>
        /// 教師編號
        /// </summary>
        public string TeacherID { get; set; }

        /// <summary>
        /// 班級名稱
        /// </summary>
        public string ClassName { get; set; }

        /// <summary>
        /// 班級年級
        /// </summary>
        public string ClassGrade { get; set; }

        /// <summary>
        /// 班級科別
        /// </summary>
        public string ClassDepartment { get; set; }

        /// <summary>
        /// 班級狀態
        /// </summary>
        public string ClassStatus { get; set; }

        /// <summary>
        /// 班級排序
        /// </summary>
        public int ClassDisplayOrder { get; set; }

        /// <summary>
        /// 輔導人數
        /// </summary>
        public int CounselPeople { get; set; }

        /// <summary>
        /// 輔導人次
        /// </summary>
        public int CounselPeopleCount { get; set; }

        /// <summary>
        /// 輔導系統老師類別
        /// </summary>
        public CounselTeacherType counselTeacherType { get; set; }
    }
}
