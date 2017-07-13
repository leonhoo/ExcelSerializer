using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelSerializer
{
 public   interface IValueConverter
    {
        /// <summary>
        /// 对象的值转换成Excel的数值
        /// </summary>
        /// <param name="value">属性值</param>
        /// <param name="targetType">属性值类型</param>
        /// <param name="parameter"></param>
        /// <returns></returns>
        string serialize(object value, Type targetType, object parameter) ;
    
        /// <summary>
        /// 从Excel读取数值后转换成对象的值
        /// </summary>
        /// <param name="value">Excel值</param>
        /// <param name="targetType">Excel值类型</param>
        /// <param name="parameter">参数</param>
        /// <returns></returns>
        object deserialize(string value, Type targetType, object parameter) ;
    }
}
