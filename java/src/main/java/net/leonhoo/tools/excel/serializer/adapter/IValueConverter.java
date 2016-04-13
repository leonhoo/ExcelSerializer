package net.leonhoo.tools.excel.serializer.adapter;

import java.lang.reflect.Type;

/**
 * 值转换器
 * 
 * @author leon
 *
 */
public interface IValueConverter {
	/**
	 * 对象的值转换成Excel的数值
	 * 
	 * @param value
	 *            属性值
	 * @param targetType
	 *            属性值类型
	 * @param parameter
	 *            参数
	 * @return
	 */
	Object serialize(Object value, Type targetType, Object parameter) throws Exception;

	/**
	 * 从Excel读取数值后转换成对象的值
	 * 
	 * @param value
	 *            Excel值
	 * @param targetType
	 *            Excel值类型
	 * @param parameter
	 *            参数
	 * @return
	 */
	Object deserialize(Object value, Type targetType, Object parameter) throws Exception;
}
