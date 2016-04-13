package net.leonhoo.tools.excel.serializer.adapter;

import java.lang.reflect.Type;

/**
 * 根据返回值自定义颜色
 * 
 * @author leon
 *
 */
public interface IColorPicker {
	/**
	 * 根据值来设置输出的颜色
	 * 
	 * @param value
	 *            对象的值
	 * @param targetType
	 *            对象的类型
	 * @param parameter
	 * @return 颜色,如"#123456"的字符串
	 * @throws Exception
	 */
	String get(Object value, Type targetType, Object parameter) throws Exception;
}
