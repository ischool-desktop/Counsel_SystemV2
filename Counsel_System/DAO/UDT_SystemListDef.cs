using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace Counsel_System2.DAO
{
    [FISCA.UDT.TableName("counsel.system_list")]
    public class UDT_SystemListDef:ActiveRecord
    {
        [FISCA.UDT.Field(Field = "name", Indexed = false)]
        public String Name { get; set; }

        [FISCA.UDT.Field(Field = "content", Indexed = false)]
        public String Content { get; set; }

    }
}
