using Npoi.Core.SS.UserModel;
using System;

namespace ExcelSerializer.Attributes
{
    [AttributeUsage(AttributeTargets.Property, Inherited = false)]
    public class ExcelAlignAttribute : Attribute
    {
        /// <summary>
        /// 值在Excel的Cell中的位置
        /// </summary>
        public HorizontalAlignment Align { get; set; }

        public ExcelAlignAttribute(HorizontalAlignment align)
        {
            this.Align = align;
        }
    }
}
