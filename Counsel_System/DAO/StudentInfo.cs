using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// 學生資訊
    /// </summary>
    public class StudentInfo
    {
        /// <summary>
        /// 學生編號
        /// </summary>
        public string StudentID { get; set; }

        /// <summary>
        /// 姓名
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 學號
        /// </summary>
        public string StudentNumber { get; set; }

        /// <summary>
        /// 身分證號
        /// </summary>
        public string IDNumber { get; set; }

        /// <summary>
        /// 入學年
        /// </summary>
        public string BeforeEnrollSchoolYear { get; set; }

        /// <summary>
        /// 入學學校名稱
        /// </summary>
        public string BeforeEnrollSchoolName { get; set; }

        /// <summary>
        /// 畢業年月
        /// </summary>
        public string GraduationYearMonth { get; set; }

        /// <summary>
        /// 出生地
        /// </summary>
        public string Birthplace { get; set; }

        /// <summary>
        /// 出生日期
        /// </summary>
        public string Birthday { get; set; }

        /// <summary>
        /// 出生日期(民國)
        /// </summary>
        public string BirthdayTW { get; set; }

        /// <summary>
        /// 戶籍地址
        /// </summary>
        public string PermanentAddress { get; set; }

        /// <summary>
        /// 戶籍電話
        /// </summary>
        public string PermanentPhone { get; set; }

        /// <summary>
        /// 聯絡地址
        /// </summary>
        public string MailingAddress { get; set; }

        /// <summary>
        /// 聯絡電話
        /// </summary>
        public string MailingPhone { get; set; }

        /// <summary>
        /// 異動紀錄字串組
        /// </summary>
        public List<string> UpdateRecordList { get; set; }

        /// <summary>
        /// 缺曠紀錄 Dictionary 年級+學期,string
        /// </summary>
        public Dictionary<string, string> AttendanceDict { get; set; }

        /// <summary>
        /// 獎懲紀錄 Dictionary 年級+學期,string
        /// </summary>
        public Dictionary<string, string> MeritDict { get; set; }

        /// <summary>
        /// 學校名稱
        /// </summary>
        public string SchoolName { get; set; }

        /// <summary>
        /// 性別
        /// </summary>
        public string Gender { get; set; }
    
    }
}
