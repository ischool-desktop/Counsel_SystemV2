using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// AB 表學生日常生活表現(導師評語)
    /// </summary>
    public class ABCard_StudentTextScore
    {
        /// <summary>
        /// 學年度
        /// </summary>
        public int SchoolYear { get; set; }

        /// <summary>
        /// 學期
        /// </summary>
        public int Semester { get; set; }

        /// <summary>
        /// 年級
        /// </summary>
        public int GradeYear { get; set; }

        /// <summary>
        /// 學生系統編號
        /// </summary>
        public string StudentID { get; set; }
        
        /// <summary>
        /// 高中(導師評語)
        /// </summary>
        public string sb_Comment { get; set; }

        /// <summary>
        /// 綜合評語(日常生活表現具體建議)
        /// </summary>
        public string DailyLifeRecommend { get; set; }
    }
}
