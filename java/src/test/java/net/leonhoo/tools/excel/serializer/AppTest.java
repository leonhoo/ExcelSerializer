package net.leonhoo.tools.excel.serializer;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

import org.junit.Test;

import net.leonhoo.tools.excel.serializer.models.Goods;
import net.leonhoo.tools.excel.serializer.models.House;
import net.leonhoo.tools.excel.serializer.models.Person;
import net.leonhoo.tools.excel.serializer.util.ExcelSerializer;

public class AppTest {

	@Test
	public void ImportTest() {
		try {
			String uri = "test.xls";
			ExcelSerializer<Person> helper = new ExcelSerializer<Person>(Person.class);
			List<Person> a = helper.load(uri);
			System.out.println(a.size());
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@Test
	public void ImportListTest() {
		try {
			// excel中存在二维数据,如 ABC DEF GHI 这几列为一个类的3个实例对应属性的值
			String uri = "testList.xls";
			ExcelSerializer<Goods> helper = new ExcelSerializer<Goods>(Goods.class);
			Map<Integer, List<Goods>> a = helper.loadList(uri);
			System.out.println(a.size());
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@Test
	public void ExportTest() {
		try {
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

	@Test
	public void ExportListTest() {
		try {
			// TODO 期待中...

		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
