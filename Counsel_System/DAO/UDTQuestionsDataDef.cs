using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace Counsel_System2.DAO
{
    /// <summary>
    /// 輔導-題目資料
    /// </summary>
    [TableName("ischool.counsel.questions_data")]
    public class UDTQuestionsDataDef:ActiveRecord
    {

        ///<summary>
        /// 群組
        ///</summary>
        [Field(Field = "group_name", Indexed = false)]
        public string Group { get; set; }


        ///<summary>
        /// 名稱
        ///</summary>
        [Field(Field = "name", Indexed = false)]
        public string Name { get; set; }

        ///<summary>
        /// 問題類型
        ///</summary>
        [Field(Field = "question_type", Indexed = false)]
        public string QuestionType { get; set; }

        ///<summary>
        /// 控制項類型
        ///</summary>
        [Field(Field = "control_type", Indexed = false)]
        public string ControlType { get; set; }

        ///<summary>
        /// 是否能列印
        ///</summary>
        [Field(Field = "can_print", Indexed = false)]
        public bool CanPrint { get; set; }

        ///<summary>
        /// 教師是否能修改
        ///</summary>
        [Field(Field = "can_teacher_edit", Indexed = false)]
        public bool CanTeacherEdit { get; set; }

        ///<summary>
        /// 學生是否能修改
        ///</summary>
        [Field(Field = "can_student_edit", Indexed = false)]
        public bool CanStudentEdit { get; set; }

        ///<summary>
        /// 顯示順序
        ///</summary>
        [Field(Field = "display_order", Indexed = false)]
        public int? DisplayOrder { get; set; }

        ///<summary>
        /// 項目內容
        ///</summary>
        [Field(Field = "items", Indexed = false)]
        public string Items { get; set; }

    }
}
