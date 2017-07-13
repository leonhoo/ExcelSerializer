using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace ExcelSerializer.Attributes
{
    [AttributeUsage(AttributeTargets.Property, Inherited = false)]
    public sealed class ExcelColumnAttribute : Attribute
    {
        public int Index { get; set; }

        public ExcelColumnAttribute(int index)
        {
            this.Index = index;
        }
    }
}
