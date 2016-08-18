using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Counsel_System2.DAO
{
    [FISCA.UDT.TableName("counsel.ab_card.data")]
    public class UDT_ABCardDataDef : FISCA.UDT.ActiveRecord
    {
        [FISCA.UDT.Field(Field = "subject_name", Indexed = false)]
        public String SubjectName { get; set; }

        [FISCA.UDT.Field(Field = "ref_template_id", Indexed = true)]
        public int TemplateID { get; set; }

        [FISCA.UDT.Field(Field = "ref_student_id", Indexed = true)]
        public int StudentID { get; set; }

        [FISCA.UDT.Field(Field = "content", Indexed = false)]
        public String Content { get; set; }

    }
}
