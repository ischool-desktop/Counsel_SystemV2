using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// 輔導類別資料
    /// </summary>
    public class CounselTeacherInterviewWaysRecord
    {
        
        /// <summary>
        /// 輔導老師名稱
        /// </summary>
        public string CounselTeacherName { get; set; }

        /// <summary>
        /// 方式以及其人數(list.count)
        /// </summary>
        public Dictionary<string, List<string>> WaysStaticsPeopleDict { get; set; }

        /// <summary>
        /// 方式以及其人次(int)
        /// </summary>
        public Dictionary<string, int> WaysStaticsPeopleCountDict { get; set; }


    }
}
