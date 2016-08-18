using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// 輔導晤談紀錄報表使用
    /// </summary>
    public class Rpt_InterviewRecord
    {
        public string StudentID { get; set; }

        public string TeacherID { get; set; }

        /// <summary>
        /// 班級
        /// </summary>
        public string ClassName { get; set; }

        /// <summary>
        /// 座號
        /// </summary>
        public int? SeatNo { get; set; }

        /// <summary>
        /// 姓名
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 教師姓名
        /// </summary>
        public string TeacherName { get; set; }

        /// <summary>
        /// 晤談方式
        /// </summary>
        public string interview_type { get; set; }

        /// <summary>
        /// 晤談對象
        /// </summary>
        public string interviewee_type { get; set; }

        /// <summary>
        /// 晤談日期
        /// </summary>
        public DateTime interview_date { get; set; }

        /// <summary>
        /// 內容要點
        /// </summary>
        public string content_digest { get; set; }

    }
}
