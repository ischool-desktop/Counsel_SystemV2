using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace Psychological_Test_Import_whsh.DAO
{

    // 文華輔導 新編多元性向測驗文華高中測驗結果 UDT
    [TableName("counsel.student_aptitude_test_data")]

    class UDT_Aptitude_Test_Data_Def : ActiveRecord
    {

        /// <summary>
        /// 語文推理_原始分數
        /// </summary>        
        [Field(Field = "language_reasoning_original_score", Indexed = false)]
        public int language_reasoning_original_score { get; set; }

        /// <summary>
        /// 語文推理_量表分數
        /// </summary>        
        [Field(Field = "language_reasoning_scale_score", Indexed = false)]
        public int language_reasoning_scale_score { get; set; }

        /// <summary>
        /// 語文推理_百分等級
        /// </summary>        
        [Field(Field = "language_reasoning_pr_level", Indexed = false)]
        public int language_reasoning_pr_level { get; set; }

        /// <summary>
        /// 數字推理_原始分數
        /// </summary>        
        [Field(Field = "digit_reasoning_original_score", Indexed = false)]
        public int digit_reasoning_original_score { get; set; }

        /// <summary>
        /// 數字推理_量表分數
        /// </summary>        
        [Field(Field = "digit_reasoning_scale_score", Indexed = false)]
        public int digit_reasoning_scale_score { get; set; }

        /// <summary>
        /// 數字推理_百分等級
        /// </summary>        
        [Field(Field = "digit_reasoning_pr_level", Indexed = false)]
        public int digit_reasoning_pr_level { get; set; }

        /// <summary>
        /// 圖形推理_原始分數
        /// </summary>        
        [Field(Field = "image_reasoning_original_score", Indexed = false)]
        public int image_reasoning_original_score { get; set; }

        /// <summary>
        /// 圖形推理_量表分數
        /// </summary>        
        [Field(Field = "image_reasoning_scale_score", Indexed = false)]
        public int image_reasoning_scale_score { get; set; }

        /// <summary>
        /// 圖形推理_百分等級
        /// </summary>        
        [Field(Field = "image_reasoning_pr_level", Indexed = false)]
        public int image_reasoning_pr_level { get; set; }

        /// <summary>
        /// 機械推理_原始分數
        /// </summary>        
        [Field(Field = "mechanical_reasoning_original_score", Indexed = false)]
        public int mechanical_reasoning_original_score { get; set; }

        /// <summary>
        /// 機械推理_量表分數
        /// </summary>        
        [Field(Field = "mechanical_reasoning_scale_score", Indexed = false)]
        public int mechanical_reasoning_scale_score { get; set; }

        /// <summary>
        /// 機械推理_百分等級
        /// </summary>        
        [Field(Field = "mechanical_reasoning_pr_level", Indexed = false)]
        public int mechanical_reasoning_pr_level { get; set; }

        /// <summary>
        /// 空間關係_原始分數
        /// </summary>        
        [Field(Field = "dimension_relation_original_score", Indexed = false)]
        public int dimension_relation_original_score { get; set; }

        /// <summary>
        /// 空間關係_量表分數
        /// </summary>        
        [Field(Field = "dimension_relation_scale_score", Indexed = false)]
        public int dimension_relation_scale_score { get; set; }

        /// <summary>
        /// 空間關係_百分等級
        /// </summary>        
        [Field(Field = "dimension_relation_pr_level", Indexed = false)]
        public int dimension_relation_pr_level { get; set; }

        /// <summary>
        /// 中文詞語_原始分數
        /// </summary>        
        [Field(Field = "chineses_words_original_score", Indexed = false)]
        public int chineses_words_original_score { get; set; }

        /// <summary>
        /// 中文詞語_量表分數
        /// </summary>        
        [Field(Field = "chineses_words_scale_score", Indexed = false)]
        public int chineses_words_scale_score { get; set; }

        /// <summary>
        /// 中文詞語_百分等級
        /// </summary>        
        [Field(Field = "chineses_words_pr_level", Indexed = false)]
        public int chineses_words_pr_level { get; set; }

        /// <summary>
        /// 英文詞語_原始分數
        /// </summary>        
        [Field(Field = "english_words_original_score", Indexed = false)]
        public int english_words_original_score { get; set; }

        /// <summary>
        /// 英文詞語_量表分數
        /// </summary>        
        [Field(Field = "english_words_scale_score", Indexed = false)]
        public int english_words_scale_score { get; set; }

        /// <summary>
        /// 英文詞語_百分等級
        /// </summary>        
        [Field(Field = "english_words_pr_level", Indexed = false)]
        public int english_words_pr_level { get; set; }

        /// <summary>
        /// 知覺速度_原始分數
        /// </summary>        
        [Field(Field = "perception_time_original_score", Indexed = false)]
        public int perception_time_original_score { get; set; }

        /// <summary>
        /// 知覺速度_量表分數
        /// </summary>        
        [Field(Field = "perception_time_scale_score", Indexed = false)]
        public int perception_time_scale_score { get; set; }

        /// <summary>
        /// 知覺速度_百分等級
        /// </summary>        
        [Field(Field = "perception_time_pr_level", Indexed = false)]
        public int perception_time_pr_level { get; set; }

        /// <summary>
        /// 學業性向_組合分數
        /// </summary>        
        [Field(Field = "learning_aptitude_assemble_score", Indexed = false)]
        public int learning_aptitude_assemble_score { get; set; }

        /// <summary>
        /// 學業性向_百分等級
        /// </summary>        
        [Field(Field = "learning_aptitude_pr_level", Indexed = false)]
        public int learning_aptitude_pr_level { get; set; }

        /// <summary>
        /// 理工性向_組合分數
        /// </summary>        
        [Field(Field = "science_aptitude_assemble_score", Indexed = false)]
        public int science_aptitude_assemble_score { get; set; }

        /// <summary>
        /// 理工性向_百分等級
        /// </summary>        
        [Field(Field = "science_aptitude_pr_level", Indexed = false)]
        public int science_aptitude_pr_level { get; set; }

        /// <summary>
        /// 文科性向_組合分數
        /// </summary>        
        [Field(Field = "literal_aptitude_assemble_score", Indexed = false)]
        public int literal_aptitude_assemble_score { get; set; }

        /// <summary>
        /// 文科性向_百分等級
        /// </summary>        
        [Field(Field = "literal_aptitude_pr_level", Indexed = false)]
        public int literal_aptitude_pr_level { get; set; }

        /// <summary>
        /// 知覺速度_作答題數
        /// </summary>        
        [Field(Field = "perception_time_complete_quiz_count", Indexed = false)]
        public int perception_time_complete_quiz_count { get; set; }

        /// <summary>
        /// 知覺速度_答對題數
        /// </summary>        
        [Field(Field = "perception_time_correct_quiz_count", Indexed = false)]
        public int perception_time_correct_quiz_count { get; set; }



        ///<summary>
        /// 學生編號
        ///</summary>
        [Field(Field = "ref_student_id", Indexed = false)]
        public int StudentID { get; set; }

        /// <summary>
        /// 實施日期
        /// </summary>        
        [Field(Field = "implementation_date", Indexed = false)]
        public DateTime? ImplementationDate { get; set; }

        





    }
}
