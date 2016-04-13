package net.leonhoo.tools.excel.serializer;

import java.lang.reflect.Type;
import java.text.SimpleDateFormat;
import java.util.Date;

import net.leonhoo.tools.excel.serializer.adapter.IValueConverter;

public class DateConverter implements IValueConverter {

	@Override
	public Object serialize(Object value, Type targetType, Object parameter) throws Exception {
		if (value != null) {
			Date v = (Date) value;
			SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
			return sdf.format(v);
		}
		return null;
	}

	@Override
	public Object deserialize(Object value, Type targetType, Object parameter) throws Exception {
		if (value != null) {
			String v = value + "";
			SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
			return sdf.parse(v);
		}
		return null;
	}

}
