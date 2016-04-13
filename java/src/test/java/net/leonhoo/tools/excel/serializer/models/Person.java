package net.leonhoo.tools.excel.serializer.models;

import java.util.Date;

import net.leonhoo.tools.excel.serializer.AgeColorPicker;
import net.leonhoo.tools.excel.serializer.DateConverter;
import net.leonhoo.tools.excel.serializer.GenderConverter;
import net.leonhoo.tools.excel.serializer.annotation.Color;
import net.leonhoo.tools.excel.serializer.annotation.Index;
import net.leonhoo.tools.excel.serializer.annotation.Display;
import net.leonhoo.tools.excel.serializer.annotation.Converter;

@Display("个人信息")
public class Person {

	@Display("姓名")
	@Index(0)
	private String name;

	@Display("年龄")
	@Index(1)
	@Color(AgeColorPicker.class)
	private int age;

	@Display("生日")
	@Index(2)
	@Converter(DateConverter.class)
	private Date birthday;

	@Display("性别")
	@Index(3)
	@Converter(GenderConverter.class)
	private boolean isMale;

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public int getAge() {
		return age;
	}

	public void setAge(int age) {
		this.age = age;
	}

	public Date getBirthday() {
		return birthday;
	}

	public void setBirthday(Date birthday) {
		this.birthday = birthday;
	}

	public boolean isMale() {
		return isMale;
	}

	public void setMale(boolean isMale) {
		this.isMale = isMale;
	}

	public boolean getIsMale() {
		return isMale;
	}

	public void setIsMale(boolean isMale) {
		this.isMale = isMale;
	}
}
