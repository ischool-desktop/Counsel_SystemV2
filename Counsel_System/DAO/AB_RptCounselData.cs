using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// AB 表用存儲輔導資料
    /// </summary>
    public class AB_RptCounselData
    {
        /// <summary>
        /// 學生系統編號
        /// </summary>
        public string StudentID { get; set; }

        /// <summary>
        /// 對應資料
        /// </summary>
        public Dictionary<string, string> DataDict = new Dictionary<string, string>();

        /// <summary>
        /// 其它資料
        /// </summary>
        public List<string> OtherDataList = new List<string>();
    }
}
