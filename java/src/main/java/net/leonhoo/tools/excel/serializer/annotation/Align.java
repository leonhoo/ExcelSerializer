package net.leonhoo.tools.excel.serializer.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Cell中所在位置
 * 
 * @author leon
 *
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ ElementType.FIELD })
@Documented
public @interface Align {
	/**
	 * 值在Cell中的位置
	 * 
	 * @author leon
	 *
	 */
	public enum EnumCellAlign {
		LEFT, CENTER, RIGHT
	}

	EnumCellAlign value();
}
