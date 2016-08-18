using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace Counsel_System2.Parser
{
    /**
    * <Col name=”稱謂” width=”” />
    * */
    public class GridColumn
    {
        private string colName = "";
        private string width = "";

        public GridColumn(XmlElement elmColumn)
        {
            this.colName = elmColumn.GetAttribute("name");
            this.width = elmColumn.GetAttribute("width");
        }

        public string GetName()
        {
            return this.colName;
        }

        public string GetWidth()
        {
            return this.width;
        }
    }
}
