package net.leonhoo.tools.excel.serializer;

import java.lang.reflect.Type;

import net.leonhoo.tools.excel.serializer.adapter.IValueConverter;

public class GenderConverter implements IValueConverter {

	@Override
	public Object serialize(Object value, Type targetType, Object parameter) throws Exception {
		if (value != null) {
			Boolean v = (boolean) value;
			return v ? "男" : "女";
		}
		return null;
	}

	@Override
	public Object deserialize(Object value, Type targetType, Object parameter) throws Exception {
		if (value != null) {
			return value.equals("男");
		}
		return null;
	}

}
