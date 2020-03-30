package com.github.entity;

import java.io.Serializable;

/**
 *
 * 表格信息映射
 *
 */
public class TableMap implements Serializable {

	private static final long serialVersionUID = -63505086622337829L;

//	表头名称
	private String key;

//	对应的实体属性名称
	private String value;

	public String getKey() {
		return key;
	}

	public void setKey(String key) {
		this.key = key;
	}

	public String getValue() {
		return value;
	}

	public void setValue(String value) {
		this.value = value;
	}

	public TableMap() {
	}

	public TableMap(String key, String value) {
		this.key = key;
		this.value = value;
	}

	@Override
	public String toString() {
		return "TableMap{" +
				"key='" + key + '\'' +
				", value='" + value + '\'' +
				'}';
	}
}
