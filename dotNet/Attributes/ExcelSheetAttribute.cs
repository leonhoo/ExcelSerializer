using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace ExcelSerializer.Attributes
{
    [AttributeUsage(AttributeTargets.Class, Inherited = false)]
    public sealed class ExcelSheetAttribute : Attribute
    {
        public int Index { get; private set; }

        public string Name { get; private set; }

        public bool IgnoreExcelHeader { get; set; } = true;

        public ExcelSheetAttribute(int index)
        {
            this.Index = index;
        }

        public ExcelSheetAttribute(string name)
        {
            this.Name = name;
        }
    }
}
