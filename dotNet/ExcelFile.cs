using ExcelSerializer.Attributes;
using Npoi.Core.HSSF.UserModel;
using Npoi.Core.SS.UserModel;
using Npoi.Core.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace ExcelSerializer
{
    public class ExcelFile
    {
        /// <summary>
        /// 从excel文件中读取数据
        /// </summary>
        /// <param name="uri"></param>
        /// <returns></returns>
        public static List<T> Load<T>(string uri)
        {
            ExcelFile excelFile = new ExcelFile();
            List<T> result = null;
            if (File.Exists(uri))
            {
                IWorkbook book = excelFile.getWorkbook(uri);
                ExcelInfo info = new ExcelInfo(typeof(T));
                ISheet sheet = excelFile.getSheet(book, info.ClazzDisplayName, info.SheetIndex);
                if (sheet == null)
                    throw new Exception("表不存在,请检查表名");

                result = info.GetData<T>(sheet);
            }
            return result;
        }

        /// <summary>
        /// 重新生成新的Excel文件
        /// </summary>
        /// <param name="uri">Excel文件路径</param>
        /// <param name="data">需要保存的数据</param>
        public static void Save(string uri, params IList[] data)
        {
            ExcelFile helper = new ExcelFile();
            IWorkbook book = helper.getWorkbook(uri);
            foreach (var item in data)
            {
                Type t = item[0].GetType();
                ExcelInfo info = new ExcelInfo(t);
                // 创建Excel工作表 指定名称和位置
                ISheet sheet = book.CreateSheet(info.ClazzDisplayName);
                // 往工作表中添加数据
                info.SetContent(ref sheet, item);
            }
            // 写入工作表完毕，关闭流
            using (FileStream stream = new FileStream(uri, FileMode.OpenOrCreate))
            {
                book.Write(stream);
            }
        }

        private IWorkbook getWorkbook(string uri)
        {
            IWorkbook result = null;
            if (File.Exists(uri))
            {
                using (FileStream fs = new FileStream(uri, FileMode.Open))
                {
                    if (uri.EndsWith(".xls"))
                    {
                        result = new HSSFWorkbook(fs); // 获取Excel文件对象
                    }
                    else
                    {
                        result = new XSSFWorkbook(fs); // 获取Excel文件对象
                    }
                }
            }
            else
            {
                if (uri.EndsWith(".xls"))
                {
                    result = new HSSFWorkbook(); // 获取Excel文件对象
                }
                else
                {
                    result = new XSSFWorkbook(); // 获取Excel文件对象
                }
            }

            return result;
        }

        /// <summary>
        /// 从Excel中获取对应表
        /// </summary>
        /// <param name="uri"></param>
        /// <returns></returns>
        private ISheet getSheet(IWorkbook book, string clazzDisplayName, int sheetIndex = 0)
        {
            ISheet sheet = null;
            if (!string.IsNullOrWhiteSpace(clazzDisplayName))
            {
                sheet = book.GetSheet(clazzDisplayName); // 指定工作表
            }
            if (sheet == null && sheetIndex > 0)
            {
                sheet = book.GetSheetAt(sheetIndex); // 指定工作表
            }
            if (sheet == null)
            {
                sheet = book.GetSheetAt(0); // 获取文件的指定工作表 默认的第一个
            }
            return sheet;
        }
    }

    class ExcelInfo
    {
        // Export
        private Dictionary<PropertyInfo, string> mapDisplayName = new Dictionary<PropertyInfo, string>();// <属性名,显示名称>
        private Dictionary<PropertyInfo, HorizontalAlignment> mapAlign = new Dictionary<PropertyInfo, HorizontalAlignment>();// <属性名,显示位置>
        private Dictionary<int, int> mapWidth = new Dictionary<int, int>();// <属性序列,显示宽度>
        private Dictionary<PropertyInfo, int> mapColumnIndex = new Dictionary<PropertyInfo, int>();// <属性名,列序号>
        private Dictionary<PropertyInfo, IValueConverter> mapValueConverter = new Dictionary<PropertyInfo, IValueConverter>();// <属性名,值转换器实例>

        public PropertyInfo[] PropertyInfos { get; set; }
        /// <summary>
        /// 类显示名称,对应Excel的Sheet的名称
        /// </summary>
        public string ClazzDisplayName { get; set; }
        /// <summary>
        /// 对应Excel的Sheet的位置
        /// </summary>
        public int SheetIndex;
        /// <summary>
        /// 设置默认单元格位置
        /// </summary>
        public HorizontalAlignment DefaultCellAlign { get; set; } = HorizontalAlignment.Center;
        /// <summary>
        /// 设置默认列宽
        /// </summary>
        public int DefaultCellWidth { get; set; } = 10;
        /// <summary>
        /// 设置默认行高
        /// </summary>
        public int DefaultCellHeight { get; set; } = 370;
        /// <summary>
        /// 是否忽略Excel的表头,即不读取第一行
        /// </summary>
        public bool IsIgnoreHeader { get; set; } = true;

        /// <summary>
        /// 获取该类中所有注解,并放入与属性对应的Map
        /// </summary>
        public ExcelInfo(Type t)
        {
            this.PropertyInfos = t.GetTypeInfo().GetProperties();
            this.ClazzDisplayName = t.Name;
            var attr = (ExcelSheetAttribute)t.GetTypeInfo().GetCustomAttribute(typeof(ExcelSheetAttribute));
            if (attr != null)
            {
                this.SheetIndex = attr.Index;
                if (!string.IsNullOrWhiteSpace(attr.Name))
                {
                    this.ClazzDisplayName = attr.Name;
                }
                this.IsIgnoreHeader = attr.IgnoreExcelHeader;
            }
            // 取属性上的自定义特性
            foreach (PropertyInfo PropertyInfo in PropertyInfos)
            {
                var descriptionAttr = (DescriptionAttribute)PropertyInfo.GetCustomAttributes(typeof(DescriptionAttribute), false)?.FirstOrDefault();
                var alignAttr = (ExcelAlignAttribute)PropertyInfo.GetCustomAttributes(typeof(ExcelAlignAttribute), false)?.FirstOrDefault();
                var columnAttr = (ExcelColumnAttribute)PropertyInfo.GetCustomAttributes(typeof(ExcelColumnAttribute), false)?.FirstOrDefault();
                var converterAttr = (ExcelConverterAttribute)PropertyInfo.GetCustomAttributes(typeof(ExcelConverterAttribute), false)?.FirstOrDefault();

                // 获取所有标有注解的列
                if (descriptionAttr != null)
                {
                    string name = descriptionAttr.Description;
                    mapDisplayName.Add(PropertyInfo, name);
                }
                if (alignAttr != null)
                {
                    mapAlign.Add(PropertyInfo, alignAttr.Align);
                }
                if (columnAttr != null)
                {
                    int index = columnAttr.Index;
                    mapColumnIndex.Add(PropertyInfo, index);
                }
                if (converterAttr != null)
                {
                    IValueConverter converter = converterAttr.Converter.GetTypeInfo().GetConstructor(System.Type.EmptyTypes).Invoke(null) as IValueConverter;
                    if (converter != null)
                        mapValueConverter.Add(PropertyInfo, converter);
                }
            }
        }

        internal List<T> GetData<T>(ISheet sheet)
        {
            var result = new List<T>();
            int rowIndex = 0;//用于记录执行的位置,以便错误出现时提示
            int columnIndex = 0;//用于记录执行的位置,以便错误出现时提示
            if (sheet.PhysicalNumberOfRows > 0)
            {
                int limitEmptyRow = 5; // 最大允许5个连续空行（超出5行则不循环下面的数据了）
                int emptyRow = 0; // 记录连续空行的个数
                T t = default(T);
                int start = this.IsIgnoreHeader ? 1 : 0;// 表头的目录不需要，从1开始
                for (int i = start; i < sheet.PhysicalNumberOfRows; i++)
                {
                    rowIndex = i;//记录位置,以便提示错误
                    IRow row = sheet.GetRow(i);
                    // 行数
                    if (emptyRow >= limitEmptyRow)
                        break; // 最大允许连续空行

                    if (row == null || row.GetCell(0) == null || string.IsNullOrWhiteSpace(row.GetCell(0) + ""))
                    {
                        emptyRow++;
                        continue;
                    }
                    emptyRow = 0;//清空空行记录
                    t = (T)Activator.CreateInstance(typeof(T));
                    // 开始赋值
                    foreach (PropertyInfo pi in this.PropertyInfos)
                    {
                        if (this.mapColumnIndex.ContainsKey(pi))
                        {
                            int index = this.mapColumnIndex[pi];
                            columnIndex = index;//记录位置,以便提示错误
                            try
                            {
                                // 读取Excel指定index列
                                string cellValue = row.GetCell(index) + "";
                                if (!string.IsNullOrWhiteSpace(cellValue))
                                {
                                    object objValue = null;
                                    if (this.mapValueConverter.ContainsKey(pi))
                                    {
                                        IValueConverter converter = this.mapValueConverter[pi];
                                        objValue = converter.deserialize(cellValue, pi.PropertyType, null);
                                    }
                                    else
                                    {
                                        objValue = this.changeType(cellValue, pi.PropertyType);
                                    }
                                    pi.SetValue(t, objValue);
                                }
                            }
                            catch (Exception ex)
                            {
                                throw new Exception(string.Format("表:{0} -- 第{1}行,第{2}列错误,{3}", this.ClazzDisplayName, rowIndex + 1, columnIndex + 1, ex.Message));
                            }
                        }
                    }
                    result.Add(t);
                }
            }
            return result;
        }

        internal void SetContent(ref ISheet sheet, IList data)
        {
            // **************设置列宽*****************
            foreach (var item in this.mapWidth)
            {
                sheet.SetColumnWidth(item.Key, item.Value);
            }
            setHeader(ref sheet);
            setBody(ref sheet, data);
        }

        private void setHeader(ref ISheet sheet)
        {
            // 设置格式
            IRow row = sheet.CreateRow(0);//index代表多少行
            row.HeightInPoints = 25;//行高

            foreach (PropertyInfo PropertyInfo in PropertyInfos)
            {
                if (mapColumnIndex.ContainsKey(PropertyInfo))
                {
                    int index = mapColumnIndex[PropertyInfo];
                    // 显示的名称
                    string displayName = PropertyInfo.Name;
                    if (mapDisplayName.ContainsKey(PropertyInfo))
                    {
                        displayName = mapDisplayName[PropertyInfo];
                    }
                    // 显示的宽度
                    int cellWidth = DefaultCellWidth;
                    if (mapWidth.ContainsKey(index))
                    {
                        cellWidth = mapWidth[index];
                    }
                    // 生成标题栏
                    ICell cell = row.CreateCell(index);//创建列
                    cell.SetCellType(CellType.String);
                    cell.CellStyle.Alignment = HorizontalAlignment.Center;
                    cell.SetCellValue(displayName);
                }
            }
        }

        private void setBody(ref ISheet sheet, IList data)
        {
            for (int i = 0; i < data.Count; i++)
            {
                //创建行
                IRow row = sheet.CreateRow(i + 1);
                foreach (PropertyInfo pi in PropertyInfos)
                {
                    if (mapColumnIndex.ContainsKey(pi))
                    {
                        // 显示的列的序号
                        int index = mapColumnIndex[pi];
                        // 获取显示的值
                        object objValue = pi.GetValue(data[i]); // 获取值
                        row.HeightInPoints = 25;//行高
                        ICell cell = row.CreateCell(index);//创建列
                        cell.CellStyle.Alignment = HorizontalAlignment.Center;
                        // 值转换
                        if (mapValueConverter.ContainsKey(pi))
                        {
                            objValue = mapValueConverter[pi].serialize(objValue, pi.PropertyType, null);
                        }
                        // 赋值
                        setCellValue(cell, objValue);
                    }
                }
            }
        }

        /// <summary>
        /// 获取cell的数据，并设置为对应的数据类型
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private object getCellValue(ICell cell)
        {
            object value = null;
            try
            {
                if (cell.CellType != CellType.Blank)
                {
                    switch (cell.CellType)
                    {
                        case CellType.Numeric:
                            // Date comes here
                            if (DateUtil.IsCellDateFormatted(cell))
                            {
                                value = cell.DateCellValue;
                            }
                            else
                            {
                                // Numeric type
                                value = cell.NumericCellValue;
                            }
                            break;
                        case CellType.Boolean:
                            // Boolean type
                            value = cell.BooleanCellValue;
                            break;
                        case CellType.Formula:
                            value = cell.CellFormula;
                            break;
                        default:
                            // String type
                            value = cell.StringCellValue;
                            break;
                    }
                }
            }
            catch (Exception)
            {
                value = "";
            }

            return value;
        }

        /// <summary>
        /// 根据数据类型设置不同类型的cell
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="obj"></param>
        private void setCellValue(ICell cell, object obj)
        {
            if (obj != null)
            {
                var type = obj.GetType();
                if (type == typeof(string))
                {
                    cell.SetCellType(CellType.String);
                    cell.SetCellValue(obj.ToString());
                }
                else if (type == typeof(int) || type == typeof(short) || type == typeof(int?) || type == typeof(short?))
                {
                    cell.SetCellType(CellType.Numeric);
                    cell.SetCellValue(int.Parse(obj + ""));
                }
                else if (type == typeof(double) || type == typeof(float) || type == typeof(decimal)
                    || type == typeof(double?) || type == typeof(float?) || type == typeof(decimal?))
                {
                    cell.SetCellType(CellType.Numeric);
                    cell.SetCellValue(double.Parse(obj + ""));
                }
                else if (type == typeof(IRichTextString))
                {
                    cell.SetCellValue((IRichTextString)obj);
                }
                else if (type == typeof(DateTime) || type == typeof(DateTime?))
                {
                    cell.SetCellType(CellType.String);
                    cell.SetCellValue(string.Format("{o:yyyy-MM-dd HH:mm:ss}", obj));
                }
                else if (type == typeof(bool) || type == typeof(bool?))
                {
                    cell.SetCellType(CellType.Boolean);
                    cell.SetCellValue((bool)obj);
                }
                else
                {
                    cell.SetCellType(CellType.String);
                    cell.SetCellValue(obj.ToString());
                }
            }
        }
        /// <summary>
        /// 将值转换成指定类型
        /// </summary>
        /// <param name="v"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        private object changeType(string v, Type type)
        {
            Type underlyingType = Nullable.GetUnderlyingType(type);
            string tmpValue = (v + "").Trim();
            if (string.IsNullOrWhiteSpace(tmpValue))
            {
                return defaultForType(type);
            }
            return Convert.ChangeType(tmpValue, underlyingType ?? type);
        }
        /// <summary>
        /// 获取该类型的默认值, 类似于default(T);
        /// </summary>
        /// <param name="targetType"></param>
        /// <returns></returns>
        private object defaultForType(Type targetType)
        {
            return targetType.GetTypeInfo().IsValueType ? Activator.CreateInstance(targetType) : null;
        }

    }
}
