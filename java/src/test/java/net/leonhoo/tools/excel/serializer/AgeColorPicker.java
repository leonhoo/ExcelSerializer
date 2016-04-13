package net.leonhoo.tools.excel.serializer;

import java.lang.reflect.Type;

import net.leonhoo.tools.excel.serializer.adapter.IColorPicker;

public class AgeColorPicker implements IColorPicker {

	@Override
	public String get(Object value, Type targetType, Object parameter) throws Exception {
		if (value != null) {
			int age = (int) value;
			if (age > 10) {
				return "#FF00FF";
			}
		}
		return null;
	}

}
