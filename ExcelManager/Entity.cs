using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelManager
{
    public class ExcelCell
    {
        public string Name { get; set; }
        public int? Index { get; set; }
    }
    public class ReadableCell : ExcelCell
    {
        public dynamic Value { get; set; }
        public string Type { get; set; }
        public Func<string, string>? ValueFormatter { get; set; }
    }

    public class MappingCell 
    {
        public ExcelCell WritableCell { get; set; }
        public ReadableCell ReadableCell { get; set; }
    }
}
