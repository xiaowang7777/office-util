package com.wjf.entity;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import java.io.Serializable;

public class WordTable implements Serializable {

	private XWPFTable table;

	private XWPFDocument document;

	public XWPFTable getTable() {
		return table;
	}

	public void setTable(XWPFTable table) {
		this.table = table;
	}

	public XWPFDocument getDocument() {
		return document;
	}

	public void setDocument(XWPFDocument document) {
		this.document = document;
	}

	public WordTable() {
	}

	public WordTable(XWPFTable table, XWPFDocument document) {
		this.table = table;
		this.document = document;
	}
}
