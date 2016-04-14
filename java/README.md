# About excel serializer for java
Automatic conversion between Excel data and T

自动映射Excel列与实体类,简化从Excel读取数据到实体对象.(本项目基于jxl操作Excel,相关包请先行引用)

# 例子
## Excel文件数据
![Structurizr](docs/images/data1.png)
## Main函数
```java
public static void main(String[] args) {
  try {
    /*********导入Excel数据********/
    String uri = "test.xls";//需要读取的xls文件,暂不支持xlsx文件
    ExcelSerializer<Person> deserializer = new ExcelSerializer<Person>(Person.class);
    List<Person> a = deserializer.load(uri);
    System.out.println(a.size());
    
    /*********导出Excel文件********/
    String uri = "testExport.xls";
    ExcelSerializer<Person> personSerializer = new ExcelSerializer<Person>(Person.class);
    List<Person> persons = new ArrayList<>();
    Person person = new Person();
    person.setAge(10);
    person.setBirthday(new Date());
    person.setIsMale(true);
    person.setName("张三");
    persons.add(person);
    person = new Person();
    person.setAge(20);
    person.setBirthday(new Date());
    person.setIsMale(true);
    person.setName("李思");
    persons.add(person);
    person = new Person();
    person.setAge(30);
    person.setBirthday(new Date());
    person.setIsMale(true);
    person.setName("王武");
    persons.add(person);
    personSerializer.save(uri, persons, false);
    // -------追加一个表-------
    ExcelSerializer<House> houseSerializer = new ExcelSerializer<House>(House.class);
    List<House> houses = new ArrayList<House>();
    House house = new House();
    house.setName("半岛国际");
    house.setAddress("滨盛路与长河路交叉口");
    houses.add(house);
    house = new House();
    house.setName("中南公寓");
    house.setAddress("滨盛路与时代大道交叉口");
    houses.add(house);
    house = new House();
    house.setName("锦绣江南");
    house.setAddress("钱江四桥落桥处");
    houses.add(house);
    houseSerializer.save(uri, houses, 1, true);
  } catch (Exception e) {
    e.printStackTrace();
  }
}
```
## Person实体类
```java
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
```
## House实体类
```java
import net.leonhoo.tools.excel.serializer.annotation.Index;
import net.leonhoo.tools.excel.serializer.annotation.Width;
import net.leonhoo.tools.excel.serializer.annotation.Display;

@Display("房子")
public class House {

	@Display("小区名称")
	@Index(0)
	private String name;

	@Display("地址")
	@Index(1)
	@Width(30)
	private String address;

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getAddress() {
		return address;
	}

	public void setAddress(String address) {
		this.address = address;
	}

}

```
## 时间值转换类
```java
import java.lang.reflect.Type;
import java.text.SimpleDateFormat;
import java.util.Date;

import net.leonhoo.tools.excel.serializer.adapter.IValueConverter;

public class DateConverter implements IValueConverter {

	@Override
	public Object serialize(Object value, Type targetType, Object parameter) throws Exception {
		if (value != null) {
			Date v = (Date) value;
			//具体时间格式请自行定义
			SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
			return sdf.format(v);
		}
		return null;
	}

	@Override
	public Object deserialize(Object value, Type targetType, Object parameter) throws Exception {
		if (value != null) {
			String v = value + "";
			//具体时间格式请自行定义
			SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
			return sdf.parse(v);
		}
		return null;
	}
}
```
## 单元格颜色配置类
```java
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
```
## 性别值转换类
```java
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
```
