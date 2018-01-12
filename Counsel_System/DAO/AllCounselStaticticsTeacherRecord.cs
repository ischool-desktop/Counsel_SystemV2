using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// 輔導類別資料
    /// </summary>
    public class AllCounselCaseStaticticsRecord
    {
        public enum CounselTeacherType { 班導師, 輔導老師, 認輔老師, 輔導主任 }

        /// <summary>
        /// 輔導類別名稱
        /// </summary>
        public string CounselCaseName { get; set; }

        /// <summary>
        /// 科別以及其人數(list.count)
        /// </summary>
        public Dictionary<string, List<string>> CaseStaticsPeopleDict { get; set; }

        /// <summary>
        /// 科別以及其人次(int)
        /// </summary>
        public Dictionary<string, int> CaseStaticsPeopleCountDict { get; set; }


    }
}
