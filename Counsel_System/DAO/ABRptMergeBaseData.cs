using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// 綜合紀錄表基本資料
    /// </summary>
    public class ABRptMergeBaseData
    {
        /// <summary>
        /// 索引
        /// </summary>
        public string Key { get; set; }

        /// <summary>
        /// 合併型態
        /// </summary>
        public EnumRptMergeType RptMergeType { get; set; }

        /// <summary>
        /// 值
        /// </summary>
        public string Value { get; set; }
    }
}
