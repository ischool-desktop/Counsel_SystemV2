using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;
namespace Counsel_System.DAO
{
    /// <summary>
    /// 輔導-綜合記錄表-生理缺陷
    /// </summary>
    [TableName("ischool.counsel.physical_defects")]
    public class UDT_AB_PhysicalDefects
    {
        ///<summary>
        /// 學生系統編號
        ///</summary>
        [Field(Field = "ref_student_id", Indexed = false)]
        public int StudentID { get; set; }

        ///<summary>
        /// 缺陷名稱
        ///</summary>
        [Field(Field = "defect", Indexed = true)]
        public string Defect { get; set; }


    }
}
