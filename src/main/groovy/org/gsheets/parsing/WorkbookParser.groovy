package org.gsheets.parsing

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook

/**
 * Provides basic support for parsing a grid of data from xml or non-xml spreadsheets.
 * 
 * @author Ken Krebs 
 */
class WorkbookParser {

	Workbook workbook
	
	int startRowIndex
	
	Map columnMap = [:]
	
	/**
	 * Constructs a WorkbookParser.
	 * 
	 * @param workbook to be parsed
	 */
	WorkbookParser(Workbook workbook) {
		assert workbook
		
		this.workbook = workbook
	}
	
	/**
	 * Parses a grid of data from the provided Workbook.
	 * 
	 * @param closure declares a parsing strategy
	 * 
	 * @return a List of data Maps
	 */
	List<Map> grid(Closure closure) {
		assert closure
		
		closure.delegate = this
		closure.call()
		
		data(workbook)
	}
	
	void headerRows(int rows) { startRowIndex = rows }
	
	void columns(Map columns) { columnMap = columns }
	
	private List<Map> data(Workbook workbook) {
		List data = []
		Sheet sheet = workbook.getSheetAt(0)
		int rows = sheet.physicalNumberOfRows - startRowIndex
		rows.times {
			Row row = sheet.getRow(it + startRowIndex)
			if (row) {
				Map rowData = rowData(row)
				data << rowData
			}
		}
		data
	}
	
	private Map rowData(Row row) {
		Map data = [:]
		columnMap.eachWithIndex { name, type, index ->
			Cell cell = row.getCell(index)
			def cellData
			if (type == String) { cellData = cellAsString(cell) }
			else if (type == BigDecimal) { cellData = cellAsBigDecimal(cell) }
			else if (type == Double) { cellData = cellAsDouble(cell) }
			else if (type == Float) { cellData = cellAsFloat(cell) }
			else if (type == Long) { cellData = cellAsLong(cell) }
			else if (type == Integer) { cellData = cellAsInteger(cell) }
			else if (type == Boolean) { cellData = cellAsBoolean(cell) }
			else if (type == Date) { cellData = cellAsDate(cell) }
			data[name] = cellData
		}
		data
	}
	
	private String cellAsString(Cell cell) { cell.toString() }
	
	private Boolean cellAsBoolean(Cell cell) { cell.booleanCellValue }
	
	private BigDecimal cellAsBigDecimal(Cell cell) { new BigDecimal(cell.toString()) }

	private Double cellAsDouble(Cell cell) { cell.numericCellValue}
	
	private Float cellAsFloat(Cell cell) { cellAsDouble(cell).floatValue() }
	
	private Long cellAsLong(Cell cell) { cellAsBigDecimal(cell).longValue() }
	
	private Integer cellAsInteger(Cell cell) { cellAsBigDecimal(cell).intValue() }
	
	private Date cellAsDate(Cell cell) { cell.dateCellValue }
	
}
