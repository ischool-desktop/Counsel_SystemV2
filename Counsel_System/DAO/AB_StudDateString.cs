using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// AB表使用處理學生日期與字串樣式
    /// </summary>
    public class AB_StudDateString
    {
        /// <summary>
        /// 學生系統編號
        /// </summary>
        public string StudentID { get; set; }

        /// <summary>
        /// 日期
        /// </summary>
        public DateTime Date { get; set; }

        /// <summary>
        /// 文字1
        /// </summary>
        public string Text1 { get; set; }

        /// <summary>
        /// 文字2
        /// </summary>
        public string Text2 { get; set; }
    }
}
