using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// AB表學期領域成績
    /// </summary>
    public class AB_SemsDomainScore
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
        /// 領域成績
        /// </summary>
        public string DomainName { get; set; }

        /// <summary>
        /// 語文領域科目成績
        /// </summary>
        public Dictionary<string, decimal?> LangSubjDict = new Dictionary<string, decimal?>();

        /// <summary>
        /// 領域成績
        /// </summary>
        public decimal? DomainScore { get; set; }

        /// <summary>
        /// 領域群組(報表用)
        /// </summary>
        public string GroupName { get; set; }

    }
}
