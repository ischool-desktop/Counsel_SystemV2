using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// 學生學期領域成績
    /// </summary>
    public class AB_StudSemsDomainScore
    {
        /// <summary>
        /// 學生系統編號
        /// </summary>
        public string StudentID { get; set; }

        /// <summary>
        /// 領域名稱對照(報表用)
        /// </summary>
        public Dictionary<string, List<string>> DomainNameDict = new Dictionary<string, List<string>>();

        /// <summary>
        /// 年級學期索引 (報表用)
        /// </summary>
        public List<AB_RptColIdx> GradeYearList = new List<AB_RptColIdx>();

        /// <summary>
        /// 學期領域成績
        /// </summary>
        public List<AB_SemsDomainScore> SemsDomainScoreList = new List<AB_SemsDomainScore>();

        /// <summary>
        /// 領域畢業成績
        /// </summary>
        public Dictionary<string, decimal?> GraduateScoreDict = new Dictionary<string, decimal?>();





    }
}
