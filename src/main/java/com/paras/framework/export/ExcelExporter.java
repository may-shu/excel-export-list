package com.paras.framework.export;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.List;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.paras.framework.export.base.ExportConfiguration;
import com.paras.framework.export.base.FieldToColumnMapping;

/**
 * Exporter utility of class to write data into the excel file.
 * @author Gaurav
 *
 * @param <T> Data Model Class to be written into excel.
 */
public class ExcelExporter {
	
	private static Logger LOGGER = Logger.getLogger( ExcelExporter.class );
	
	private static final int FIRST_ROW = 0;

	@SuppressWarnings("rawtypes")
	public void export( ExportConfiguration configuration ) {
		LOGGER.info("In ExcelExporter | Starting Execution of export" );
		
		String name = configuration.getExcelName();
		FieldToColumnMapping[] mappings = configuration.getMappings();
		
		int mappingsCount = mappings.length,i;
		
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet( name );
		
		/* Creating Header Row. */
		Row header = sheet.createRow( FIRST_ROW );
		
		for( i=0; i < mappingsCount; i++ ) {
			Cell cell = header.createCell( i );
			cell.setCellValue( mappings[i].getColumn() );
		}
		
		/* Writing Values */
		int currRow = 2;
		
		List data = configuration.getData();
		
		for( Object t : data ) {
			Row row = sheet.createRow( currRow++ );
			
			for( i=0; i < mappingsCount; i++ ) {
				Cell cell = row.createCell( i );
				
				Method method = getGetterMethod(configuration, mappings[i].getField());
				
				try {		
					Object result = method.invoke( t );
					
					if( result == null ) {
						cell.setCellValue( "" );
					}
					if( result instanceof Long ) {
						cell.setCellValue( (Long) result );
					} else {
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
		
		try{
			FileOutputStream out = new FileOutputStream( "C:\\Users\\Gaurav\\Desktop\\" + name + ".xlsx" );
			
			workbook.write( out );
			out.close();
			workbook.close();
		} catch( IOException ex ) {
			LOGGER.error( "In ExcelExporter | Caught IOException " + ex.getMessage() );
		}
		
		LOGGER.info("In ExcelExporter | Finished Execution of export" );
	}
	
	@SuppressWarnings("rawtypes")
	private Method getGetterMethod( ExportConfiguration configuration, String field ) {
		Class type = configuration.getType();
		
		Method[] allMethods = type.getMethods();
		String getterName = convertToGetter( field );
		
		for( Method method : allMethods ) {
			if( method.getName().equals( getterName )) {
				return method;
			}
		}
		
		return null;
	}
	
	private static String convertToGetter( String field ) {
		return "get" + String.valueOf( field.charAt( FIRST_ROW )).toUpperCase() + field.substring(1);
	}
}
