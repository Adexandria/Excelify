using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelify.Models
{
    public interface IExcelifySheet
    {
        public Stream File { get; set; }
    }
}
