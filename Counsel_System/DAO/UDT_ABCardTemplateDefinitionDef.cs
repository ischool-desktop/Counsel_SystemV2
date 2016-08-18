using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;
namespace Counsel_System2.DAO
{
    [TableName("counsel.ab_card.template")]
    public class UDT_ABCardTemplateDefinitionDef : FISCA.UDT.ActiveRecord
    {
        [Field(Field = "subject_name", Indexed = false)]
        public String SubjectName { get; set; }

        [Field(Field = "content", Indexed = false)]
        public String Content { get; set; }

        [Field(Field = "priority", Indexed = false)]
        public int? Priority { get; set; }

        public override string ToString()
        {
            return this.SubjectName;
        }
    }
}
