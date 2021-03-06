package com.paras.framework.export;

import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import org.apache.commons.io.IOUtils;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;

import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.security.crypto.codec.Base64;

import com.paras.framework.export.base.ExportConfiguration;
import com.paras.framework.export.base.FieldToColumnMapping;

/**
 * Exporter utility of class to write data into the excel file.
 * @author Gaurav
 *
 */
public class ExcelExporter {
	
	private static Logger LOGGER = Logger.getLogger( ExcelExporter.class );
	
	private static final int FIRST_ROW = 0;
	
	private static final SimpleDateFormat dateFormat = new SimpleDateFormat("MMM dd, yyyy");
	private static final SimpleDateFormat hourFormat = new SimpleDateFormat( "hh:mm:ss" );
	private static final SimpleDateFormat dateAndHourFormat = new SimpleDateFormat( "MMM dd, yyyy hh:mm" );
	
	/**
	 * Utility method to check to see if date has time part.
	 */
	private boolean hasTimePart( Date date ) {
		return hourFormat.format( date ).equals("00:00:00");
	}

	/**
	 * In this method we will try to create a temporary file.
	 * Write to it, and return its base64 encoded string.
	 * It can get ugly. :(
	 * 
	 * @param configuration
	 */
	public String export( ExportConfiguration configuration ) {
		LOGGER.info("In ExcelExporter | Starting Execution of export" );
		
		String name = configuration.getExcelName(), base64=null;
		FieldToColumnMapping[] mappings = configuration.getMappings();
		
		int mappingsCount = mappings.length,i;
		
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet( name );
		
		/* Creating Header Row. */
		Row header = sheet.createRow( FIRST_ROW );
		
		/* Creating Header Background color. */
		XSSFColor headerColor = new XSSFColor( new Color( 155, 194, 230 ));
		CellStyle headerStyle = workbook.createCellStyle();
		
		headerStyle.setFillForegroundColor( headerColor.getIndex() );
		
		for( i=0; i < mappingsCount; i++ ) {
			Cell cell = header.createCell( i );
			
			cell.setCellValue( mappings[i].getColumn() );
			cell.setCellStyle( headerStyle );
		}
		
		/* Writing Values */
		int currRow = 2;
		
		List<Object> data = configuration.getData();
		
		for( Object t : data ) {
			Row row = sheet.createRow( currRow++ );
			
			for( i=0; i < mappingsCount; i++ ) {
				Cell cell = row.createCell( i );
				
				String field = mappings[i].getField();
				
				LOGGER.debug( "In ExcelExport | export | Trying to fetch method for " + field );
				Method method = getGetterMethod(configuration, field);
				
				try {		
					Object result = method.invoke( t );
					
					if( result == null ) {
						cell.setCellValue( "" );
					}
					if( result instanceof Long ) {
						cell.setCellValue( (Long) result );
					} 
					else if( result instanceof Boolean ) {
						cell.setCellValue( (Boolean) result );
					}
					else if( result instanceof Date ) {
						
						cell.setCellType( XSSFCell.CELL_TYPE_NUMERIC );
						
						Date value = (Date) result;
						SimpleDateFormat format = null;
						
						if( hasTimePart( value )) {
							format = dateAndHourFormat;
						} else {
							format = dateFormat;
						}
						
						cell.setCellValue( format.format( value ));						
						
					}
					else {
						cell.setCellValue( (String) result );
					}
					
				} catch( IllegalAccessException ex ) {
					LOGGER.error( "In ExcelExporter | Caught IllegalAccessException " + ex.getMessage() );					
				} catch( IllegalArgumentException ex ) {
					LOGGER.error( "In ExcelExporter | Caught IllegalArgumentException " + ex.getMessage() );
				} catch( InvocationTargetException ex ) {
					LOGGER.error( "In ExcelExporter | Caught InvocationTargetException " + ex.getMessage() );
				}
			}
		}
		
		/* Setting Auto Resize. */
		for( i=0; i < mappingsCount; i++ ) {
			sheet.autoSizeColumn(i);
		}
		
		try{
			File temp = File.createTempFile( name, ".xlsx" );
			temp.deleteOnExit();
			
			FileOutputStream out = new FileOutputStream( temp );
			
			workbook.write( out );
			out.close();
			workbook.close();
			
			base64 = new String( Base64.encode( IOUtils.toByteArray( new FileInputStream( temp ))));
			
		} catch( IOException ex ) {
			LOGGER.error( "In ExcelExporter | Caught IOException " + ex.getMessage() );
		}
		
		LOGGER.info("In ExcelExporter | Finished Execution of export" );
		
		return base64;
	}
	
	/**
	 * Utility method to compute getter of a method.
	 * If the property name is id, then getter should be getId().
	 * So, this method will return a method instance whose name would be getId.
	 * 
	 * @param configuration
	 * @param field
	 * @return
	 */
	private Method getGetterMethod( ExportConfiguration configuration, String field ) {
		
		@SuppressWarnings("rawtypes")
		Class type = configuration.getType();
		
		Method[] allMethods = type.getMethods();
		String getterName = convertToGetter( field );
		
		for( Method method : allMethods ) {
			if( method.getName().equals( getterName )) {
				return method;
			}
		}
		
		LOGGER.error( "In ExcelExport | export | No suitable getter found for " + field );
		LOGGER.info( "In ExcelExport | export | Having another go considering " + field + " a boolean variable " );
		getterName = convertToGetterForBoolean( field );
		
		for( Method method : allMethods ) {
			if( method.getName().equals( getterName )) {
				return method;
			}
		}
		
		LOGGER.error( "In ExcelExport | export | No suitable getter found for " + field + " | Returning Null." );
		
		return null;
	}
	
	/**
	 * Utility method to compute getter string of a method.
	 * @param field
	 * @return
	 */
	private static String convertToGetter( String field ) {
		return "get" + String.valueOf( field.charAt( FIRST_ROW )).toUpperCase() + field.substring(1);
	}
	
	/**
	 * Utility method to compute getter string for a boolean variable.
	 */
	private static String convertToGetterForBoolean( String field ) {
		return "is" + String.valueOf( field.charAt( FIRST_ROW )).toUpperCase() + field.substring(1);
	}
}
