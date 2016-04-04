package com.paras.framework.export.base;

/**
 * Class field to column Name mapping.
 * @author Gaurav
 *
 */
public class FieldToColumnMapping {
	
	/**
	 * Field in the model.
	 */
	private String field;
	
	/**
	 * Column Name in the excel.
	 */
	private String column;

	public String getField() {
		return field;
	}

	public void setField(String field) {
		this.field = field;
	}

	public String getColumn() {
		return column;
	}

	public void setColumn(String column) {
		this.column = column;
	}
	
	public FieldToColumnMapping( String field, String column ) {
		this.field = field;
		this.column = column;
	}
}
