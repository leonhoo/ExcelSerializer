using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelSerializer.Attributes
{
    [AttributeUsage(AttributeTargets.Property, Inherited = false)]
    public class ExcelConverterAttribute : Attribute
    {
        /// <summary>
        /// 值在Excel的Cell中的颜色
        /// </summary>
        public Type Converter { get; set; }

        public ExcelConverterAttribute(Type converter)
        {
            this.Converter = converter;
        }
    }
}
