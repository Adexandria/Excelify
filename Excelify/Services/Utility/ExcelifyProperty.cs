using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelify.Services.Utility
{
    internal class ExcelifyProperty(string name, object attributeName, object type )
    {
        public string PropertyName { get; set; } = name;

        public object PropertyType = type;
        public object AttributeName { get; set; } = attributeName;
    }
}
