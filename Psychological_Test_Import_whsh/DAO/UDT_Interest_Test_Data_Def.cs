﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace Psychological_Test_Import_whsh.DAO
{

    // 文華輔導 興趣測驗結果 UDT
    [TableName("counsel.student_interest_test_data")]

    class UDT_Interest_Test_Data_Def : ActiveRecord
    {

        /// <summary>
        /// 學校代碼
        /// </summary>        
        [Field(Field = "school_code", Indexed = false)]
        public string school_code { get; set; }

        /// <summary>
        /// 班級代碼
        /// </summary>        
        [Field(Field = "class_code", Indexed = false)]
        public string class_code { get; set; }

        /// <summary>
        /// 座號
        /// </summary>        
        [Field(Field = "seat_number", Indexed = false)]
        public string seat_number { get; set; }

        /// <summary>
        /// 身分證號
        /// </summary>        
        [Field(Field = "id_card_number", Indexed = false)]
        public string id_card_number { get; set; }

        /// <summary>
        /// 抓週第一碼
        /// </summary>        
        [Field(Field = "gift_test_first_code", Indexed = false)]
        public string gift_test_first_code { get; set; }

        /// <summary>
        /// 抓週第二碼
        /// </summary>        
        [Field(Field = "gift_test_second_code", Indexed = false)]
        public string gift_test_second_code { get; set; }

        /// <summary>
        /// 抓週第三碼
        /// </summary>        
        [Field(Field = "gift_test_third_code", Indexed = false)]
        public string gift_test_third_code { get; set; }

        /// <summary>
        /// 興趣第一碼
        /// </summary>        
        [Field(Field = "interest_first_code", Indexed = false)]
        public string interest_first_code { get; set; }

        /// <summary>
        /// 興趣第二碼
        /// </summary>        
        [Field(Field = "interest_second_code", Indexed = false)]
        public string interest_second_code { get; set; }

        /// <summary>
        /// 興趣第三碼
        /// </summary>        
        [Field(Field = "interest_third_code", Indexed = false)]
        public string interest_third_code { get; set; }

        /// <summary>
        /// R型總分
        /// </summary>        
        [Field(Field = "r_type_score", Indexed = false)]
        public string r_type_score { get; set; }

        /// <summary>
        /// I型總分
        /// </summary>        
        [Field(Field = "i_type_score", Indexed = false)]
        public string i_type_score { get; set; }

        /// <summary>
        /// A型總分
        /// </summary>        
        [Field(Field = "a_type_score", Indexed = false)]
        public string a_type_score { get; set; }

        /// <summary>
        /// S型總分
        /// </summary>        
        [Field(Field = "s_type_score", Indexed = false)]
        public string s_type_score { get; set; }

        /// <summary>
        /// E型總分
        /// </summary>        
        [Field(Field = "e_type_score", Indexed = false)]
        public string e_type_score { get; set; }

        /// <summary>
        /// C型總分
        /// </summary>        
        [Field(Field = "c_type_score", Indexed = false)]
        public string c_type_score { get; set; }

        /// <summary>
        /// 興趣代碼
        /// </summary>        
        [Field(Field = "interest_code", Indexed = false)]
        public string interest_code { get; set; }

        /// <summary>
        /// 扇形區域
        /// </summary>        
        [Field(Field = "sector_area", Indexed = false)]
        public string sector_area { get; set; }

        /// <summary>
        /// 協和度
        /// </summary>        
        [Field(Field = "coordinate_index", Indexed = false)]
        public string coordinate_index { get; set; }

        /// <summary>
        /// 區分值
        /// </summary>        
        [Field(Field = "distinguishing_index", Indexed = false)]
        public string distinguishing_index { get; set; }

        /// <summary>
        /// 一致性
        /// </summary>        
        [Field(Field = "consistency_index", Indexed = false)]
        public string consistency_index { get; set; }

       
        ///<summary>
        /// 學生編號
        ///</summary>
        [Field(Field = "ref_student_id", Indexed = false)]
        public string StudentID { get; set; }

        /// <summary>
        /// 實施日期
        /// </summary>        
        [Field(Field = "implementation_date", Indexed = false)]
        public DateTime? ImplementationDate { get; set; }
        
    }
}
