package com.paras.framework.export.base;

import java.util.List;

/**
 * Excel Export Configuration.
 * This class basically stores all information how the excel is going to be written.
 * @author Gaurav
 *
 */
public class ExportConfiguration {
	/**
	 * Column Names.
	 */
	private FieldToColumnMapping[] mappings;
	
	/**
	 * Final Name of the excel file to be written.
	 */
	private String excelName;
	
	/**
	 * Generic Type.
	 */
	@SuppressWarnings("rawtypes")
	private Class type;
	
	/**
	 * List of objects to be written.
	 */
	@SuppressWarnings("rawtypes")
	private List data;

	@SuppressWarnings("rawtypes")
	public ExportConfiguration(Class type, List result, FieldToColumnMapping[] mappings, String excelName) {
		this.type = type;
		this.data = result;
		this.mappings = mappings;
		this.excelName = excelName;
	}

	public String getExcelName() {
		return excelName;
	}

	public void setExcelName(String excelName) {
		this.excelName = excelName;
	}

	@SuppressWarnings("rawtypes")
	public List getData() {
		return data;
	}

	@SuppressWarnings("rawtypes")
	public void setData(List data) {
		this.data = data;
	}

	public FieldToColumnMapping[] getMappings() {
		return mappings;
	}

	public void setMappings(FieldToColumnMapping[] mappings) {
		this.mappings = mappings;
	}

	@SuppressWarnings("rawtypes")
	public Class getType() {
		return type;
	}

	@SuppressWarnings("rawtypes")
	public void setType(Class type) {
		this.type = type;
	}
}
