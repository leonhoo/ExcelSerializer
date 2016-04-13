package net.leonhoo.tools.excel.serializer.models;

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
