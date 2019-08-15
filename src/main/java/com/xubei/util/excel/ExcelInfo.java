package com.xubei.util.excel;

import java.util.List;
import java.util.Map;

/**
 * excel操作的pojo<br>
 * titles的顺序和fields顺序保持一致
 * 
 * @author shuheng
 *
 */
public class ExcelInfo {
	/** 文件名 */
	private String fileName;
	/** sheet名 */
	private String sheetName;
	/** sheet中标题（中文） */
	private String[] titles;
	/** sheet中标题（字段） */
	private String[] fields;
	/** sheet中内容 */
	private List<Map<String, Object>> mapList;
	/** Excel每个工作簿的行数*/
	private int pageSize = 20;

	public ExcelInfo() {}

	public ExcelInfo(String fileName, String sheetName, String[] titles, String[] fields,
			List<Map<String, Object>> list) {
		this.fileName = fileName;
		this.sheetName = sheetName;
		this.titles = titles;
		this.fields = fields;
		this.mapList = list;
	}

	public String getFileName() {
		return fileName;
	}

	public void setFileName(String fileName) {
		this.fileName = fileName;
	}

	public String getSheetName() {
		return sheetName;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	public String[] getTitles() {
		return titles;
	}

	public void setTitles(String[] titles) {
		this.titles = titles;
	}

	public String[] getFields() {
		return fields;
	}

	public void setFields(String[] fields) {
		this.fields = fields;
	}

	public List<Map<String, Object>> getMapList() {
		return mapList;
	}

	public void setMapList(List<Map<String, Object>> mapList) {
		this.mapList = mapList;
	}

	public int getPageSize() {
		return pageSize;
	}

	public void setPageSize(int pageSize) {
		this.pageSize = pageSize;
	}

}
