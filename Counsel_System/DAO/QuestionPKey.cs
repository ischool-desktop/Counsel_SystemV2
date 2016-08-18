using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// 題組Key 組合
    /// </summary>
    public class QuestionPKey
    {
        /// <summary>
        /// 題組名稱
        /// </summary>
        public string SubjectLabel { get; set; }

        /// <summary>
        /// 群組名稱
        /// </summary>
        public string QGLabel { get; set; }
        
        /// <summary>
        /// 問題名稱
        /// </summary>
        public string QLabel { get; set; }

        /// <summary>
        /// 樣版ID
        /// </summary>
        public int TemplateID { get; set; }

        /// <summary>
        /// 問題類型
        /// </summary>
        public string QType { get; set; }

        /// <summary>
        /// 問題系統內 Key,不能重複值
        /// </summary>
        public string Q_Name { get; set; }

        /// <summary>
        /// 題目選項，用;分隔
        /// </summary>
        public string Q_Choice { get; set; }
        

        /// <summary>
        /// 取得 key 組
        /// </summary>
        /// <returns></returns>
        public string GetPKey()
        {
            return SubjectLabel + "_" + QGLabel + "_" + QLabel + "_" + QType;
        }

        /// <summary>
        /// 取得欄位(SubjectLabel+QGLabel+QLabel
        /// </summary>
        /// <returns></returns>
        public string GetField()
        {
            // 因為合併列印法處理數字開頭與點
            string str=("C_" + SubjectLabel + "_" + QGLabel + "_" + QLabel).Replace(".", "_");            

            return str;
        }

        /// <summary>
        /// 取得欄位(SubjectLabel+QGLabel
        /// </summary>
        /// <returns></returns>
        public string GetFieldG()
        {
            // 因為合併列印法處理數字開頭與點
            string str = ("G_" + SubjectLabel + "_" + QGLabel ).Replace(".", "_");

            if ((SubjectLabel + QGLabel) == (SubjectLabel + QGLabel + QLabel))
                str = GetField();

            return str;
        }


        /// <summary>
        /// 問題結構
        /// </summary>
        public XElement dataElement { get; set; }
    }
}
