package net.leonhoo.tools.excel.serializer.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import net.leonhoo.tools.excel.serializer.adapter.IValueConverter;

/**
 * 值转换器
 * 
 * @author leon
 *
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ ElementType.FIELD })
@Documented
public @interface Converter {
	Class<? extends IValueConverter> value();
}
