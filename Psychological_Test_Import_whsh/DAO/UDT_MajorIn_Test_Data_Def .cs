using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace Psychological_Test_Import_whsh.DAO
{

    // 文華輔導 匯入學系探索量表測驗 UDT
    [TableName("counsel.student_majorin_test_data")]

    class UDT_MajorIn_Test_Data_Def : ActiveRecord
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
        /// 學生姓名
        /// </summary>        
        [Field(Field = "student_name", Indexed = false)]
        public string student_name { get; set; }


        /// <summary>
        /// 學生座號
        /// </summary>        
        [Field(Field = "seat_number", Indexed = false)]
        public string seat_number { get; set; }

        /// <summary>
        /// 身分證號
        /// </summary>        
        [Field(Field = "id_card_number", Indexed = false)]
        public string id_card_number { get; set; }

        /// <summary>
        /// 志願一
        /// </summary>        
        [Field(Field = "wish_1", Indexed = false)]
        public string wish_1 { get; set; }

        /// <summary>
        /// 志願二
        /// </summary>        
        [Field(Field = "wish_2", Indexed = false)]
        public string wish_2 { get; set; }

        /// <summary>
        /// 志願三
        /// </summary>        
        [Field(Field = "wish_3", Indexed = false)]
        public string wish_3 { get; set; }

        /// <summary>
        /// 志願四
        /// </summary>        
        [Field(Field = "wish_4", Indexed = false)]
        public string wish_4 { get; set; }

        /// <summary>
        /// 志願五
        /// </summary>        
        [Field(Field = "wish_5", Indexed = false)]
        public string wish_5 { get; set; }

        /// <summary>
        /// 志願六
        /// </summary>        
        [Field(Field = "wish_6", Indexed = false)]
        public string wish_6 { get; set; }

        /// <summary>
        /// 志願七
        /// </summary>        
        [Field(Field = "wish_7", Indexed = false)]
        public string wish_7 { get; set; }

        /// <summary>
        /// 志願八
        /// </summary>        
        [Field(Field = "wish_8", Indexed = false)]
        public string wish_8 { get; set; }

        /// <summary>
        /// 數學分數
        /// </summary>        
        [Field(Field = "math_score", Indexed = false)]
        public string math_score { get; set; }

        /// <summary>
        /// 物理分數
        /// </summary>        
        [Field(Field = "phycics_score", Indexed = false)]
        public string phycics_score { get; set; }

        /// <summary>
        /// 化學分數
        /// </summary>        
        [Field(Field = "chemistry_score", Indexed = false)]
        public string chemistry_score { get; set; }

        /// <summary>
        /// 資訊電子分數
        /// </summary>        
        [Field(Field = "information_electronics_score", Indexed = false)]
        public string information_electronics_score { get; set; }

        /// <summary>
        /// 通訊電信分數
        /// </summary>        
        [Field(Field = "communication_telecommunications", Indexed = false)]
        public string communication_telecommunications { get; set; }

        /// <summary>
        /// 工程科技分數
        /// </summary>        
        [Field(Field = "engineer_technology_score", Indexed = false)]
        public string engineer_technology_score { get; set; }

        /// <summary>
        /// 機械分數
        /// </summary>        
        [Field(Field = "mechanism_score", Indexed = false)]
        public string mechanism_score { get; set; }

        /// <summary>
        /// 建築營造分數
        /// </summary>        
        [Field(Field = "building_construction_score", Indexed = false)]
        public string building_construction_score { get; set; }

        /// <summary>
        /// 設計分數
        /// </summary>        
        [Field(Field = "design_score", Indexed = false)]
        public string design_score { get; set; }

        /// <summary>
        /// 生命科學分數
        /// </summary>        
        [Field(Field = "life_science_score", Indexed = false)]
        public string life_science_score { get; set; }

        /// <summary>
        /// 醫學分數
        /// </summary>        
        [Field(Field = "medical_score", Indexed = false)]
        public string medical_score { get; set; }

        /// <summary>
        /// 生資食科分數
        /// </summary>        
        [Field(Field = "food_science_score", Indexed = false)]
        public string food_science_score { get; set; }

        /// <summary>
        /// 地球環境分數
        /// </summary>        
        [Field(Field = "earth_enviroment_score", Indexed = false)]
        public string earth_enviroment_score { get; set; }

        /// <summary>
        /// 藝術分數
        /// </summary>        
        [Field(Field = "art_score", Indexed = false)]
        public string art_score { get; set; }

        /// <summary>
        /// 歷史文化分數
        /// </summary>        
        [Field(Field = "history_score", Indexed = false)]
        public string history_score { get; set; }

        /// <summary>
        /// 傳播媒體分數
        /// </summary>        
        [Field(Field = "media_score", Indexed = false)]
        public string media_score { get; set; }

        /// <summary>
        /// 教育訓練分數
        /// </summary>        
        [Field(Field = "education_score", Indexed = false)]
        public string education_scorescore { get; set; }

        /// <summary>
        /// 心理學分數
        /// </summary>        
        [Field(Field = "psychology_score", Indexed = false)]
        public string psychology_score { get; set; }

        /// <summary>
        /// 社會人類分數
        /// </summary>        
        [Field(Field = "society_score", Indexed = false)]
        public string society_score { get; set; }

        /// <summary>
        /// 哲學宗教分數
        /// </summary>        
        [Field(Field = "philosophy_score", Indexed = false)]
        public string philosophy_score { get; set; }

        /// <summary>
        /// 治療諮商分數
        /// </summary>        
        [Field(Field = "consultation_score", Indexed = false)]
        public string consultation_score { get; set; }

        /// <summary>
        /// 語文文學分數
        /// </summary>        
        [Field(Field = "language_score", Indexed = false)]
        public string language_score { get; set; }

        /// <summary>
        /// 外國語文分數
        /// </summary>        
        [Field(Field = "foreign_language_score", Indexed = false)]
        public string foreign_language_score { get; set; }

        /// <summary>
        /// 人力資源分數
        /// </summary>        
        [Field(Field = "human_resource_score", Indexed = false)]
        public string human_resource_score { get; set; }

        /// <summary>
        /// 顧客服務分數
        /// </summary>        
        [Field(Field = "customer_service_score", Indexed = false)]
        public string customer_service_score { get; set; }

        /// <summary>
        /// 管理分數
        /// </summary>        
        [Field(Field = "management_score", Indexed = false)]
        public string management_score { get; set; }

        /// <summary>
        /// 銷售行銷分數
        /// </summary>        
        [Field(Field = "market_score", Indexed = false)]
        public string market_score { get; set; }

        /// <summary>
        /// 經濟會計分數
        /// </summary>        
        [Field(Field = "accounting_score", Indexed = false)]
        public string accounting_score { get; set; }

        /// <summary>
        /// 法律政治分數
        /// </summary>        
        [Field(Field = "law_score", Indexed = false)]
        public string law_score { get; set; }

        /// <summary>
        /// 行政分數
        /// </summary>        
        [Field(Field = "administrative_score", Indexed = false)]
        public string administrative_score { get; set; }


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
