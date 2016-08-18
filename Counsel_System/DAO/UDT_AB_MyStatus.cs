using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace Counsel_System.DAO
{
    /// <summary>
    /// 輔導-綜合記錄表-本人概況
    /// </summary>
    [TableName("ischool.counsel.my_status")]
    public class UDT_AB_MyStatus:ActiveRecord
    {
        ///<summary>
        /// 學生系統編號
        ///</summary>
        [Field(Field = "ref_student_id", Indexed = false)]
        public int StudentID { get; set; }

        ///<summary>
        /// 血型
        ///</summary>
        [Field(Field = "blood_type", Indexed = false)]
        public string BloodType { get; set; }

        ///<summary>
        /// 宗教
        ///</summary>
        [Field(Field = "religion", Indexed = false)]
        public string Religion { get; set; }

        ///<summary>
        /// 身高1上
        ///</summary>
        [Field(Field = "h_1a", Indexed = false)]
        public decimal? H1A { get; set; }

        ///<summary>
        /// 身高1下
        ///</summary>
        [Field(Field = "h_1b", Indexed = false)]
        public decimal? H1B { get; set; }

        ///<summary>
        /// 身高2上
        ///</summary>
        [Field(Field = "h_2a", Indexed = false)]
        public decimal? H2A { get; set; }

        ///<summary>
        /// 身高2下
        ///</summary>
        [Field(Field = "h_2b", Indexed = false)]
        public decimal? H2B { get; set; }

        ///<summary>
        /// 身高3上
        ///</summary>
        [Field(Field = "h_3a", Indexed = false)]
        public decimal? H3A { get; set; }

        ///<summary>
        /// 身高3下
        ///</summary>
        [Field(Field = "h_3b", Indexed = false)]
        public decimal? H3B { get; set; }

        ///<summary>
        /// 身高4上
        ///</summary>
        [Field(Field = "h_4a", Indexed = false)]
        public decimal? H4A { get; set; }

        ///<summary>
        /// 身高4下
        ///</summary>
        [Field(Field = "h_4b", Indexed = false)]
        public decimal? H4B { get; set; }

        ///<summary>
        /// 身高5上
        ///</summary>
        [Field(Field = "h_5a", Indexed = false)]
        public decimal? H5A { get; set; }

        ///<summary>
        /// 身高5下
        ///</summary>
        [Field(Field = "h_5b", Indexed = false)]
        public decimal? H5B { get; set; }

        ///<summary>
        /// 身高6上
        ///</summary>
        [Field(Field = "h_6a", Indexed = false)]
        public decimal? H6A { get; set; }

        ///<summary>
        /// 身高6下
        ///</summary>
        [Field(Field = "h_6b", Indexed = false)]
        public decimal? H6B { get; set; }

        ///<summary>
        /// 體重1上
        ///</summary>
        [Field(Field = "w_1a", Indexed = false)]
        public decimal? W1A { get; set; }

        ///<summary>
        /// 體重1下
        ///</summary>
        [Field(Field = "w_1b", Indexed = false)]
        public decimal? W1B { get; set; }

        ///<summary>
        /// 體重2上
        ///</summary>
        [Field(Field = "w_2a", Indexed = false)]
        public decimal? W2A { get; set; }

        ///<summary>
        /// 體重2下
        ///</summary>
        [Field(Field = "w_2b", Indexed = false)]
        public decimal? W2B { get; set; }

        ///<summary>
        /// 體重3上
        ///</summary>
        [Field(Field = "w_3a", Indexed = false)]
        public decimal? W3A { get; set; }

        ///<summary>
        /// 體重3下
        ///</summary>
        [Field(Field = "w_3b", Indexed = false)]
        public decimal? W3B { get; set; }

        ///<summary>
        /// 體重4上
        ///</summary>
        [Field(Field = "w_4a", Indexed = false)]
        public decimal? W4A { get; set; }

        ///<summary>
        /// 體重4下
        ///</summary>
        [Field(Field = "w_4b", Indexed = false)]
        public decimal? W4B { get; set; }

        ///<summary>
        /// 體重5上
        ///</summary>
        [Field(Field = "w_5a", Indexed = false)]
        public decimal? W5A { get; set; }

        ///<summary>
        /// 體重5下
        ///</summary>
        [Field(Field = "w_5b", Indexed = false)]
        public decimal? W5B { get; set; }

        ///<summary>
        /// 體重6上
        ///</summary>
        [Field(Field = "w_6a", Indexed = false)]
        public decimal? W6A { get; set; }


    }
}
