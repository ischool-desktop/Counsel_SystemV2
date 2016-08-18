using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// AB表用文字
    /// </summary>
    public class AB_StudText
    {
        /// <summary>
        /// 學生系統編號
        /// </summary>
        public string StudentID { get; set; }

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
        /// 文字
        /// </summary>
        public string Text1 { get; set; }
        
    }
}
