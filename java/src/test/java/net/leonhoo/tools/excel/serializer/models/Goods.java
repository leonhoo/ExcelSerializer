package net.leonhoo.tools.excel.serializer.models;

import net.leonhoo.tools.excel.serializer.annotation.Indexs;

public class Goods {
	@Indexs({ 2, 5, 8 })
	private Integer id;
	@Indexs({ 3, 6, 9 })
	private String name;
	@Indexs({ 4, 7, 10 })
	private Double price;

	public Integer getId() {
		return id;
	}

	public void setId(Integer id) {
		this.id = id;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public Double getPrice() {
		return price;
	}

	public void setPrice(Double price) {
		this.price = price;
	}

}
