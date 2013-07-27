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
	
	// TODO: add error reporting and failOnErrors property

	private final Workbook workbook
	
	private int startRowIndex
	
	private Map columnMap = [:]
	
	private final Map convertors = [
		string: this.&cellAsString,
		decimal: this.&cellAsBigDecimal,
		'int': this.&cellAsInteger,
		'boolean': this.&cellAsBoolean,
		'double': this.&cellAsDouble,
		'float': this.&cellAsFloat,
		date: this.&cellAsDate,
		'long': this.&cellAsLong,
	]
	
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
	
	/**
	 * Sets the no. of header rows to skip when parsing a grid.
	 * 
	 * @param rows
	 */
	void headerRows(int rows) { startRowIndex = rows }
	
	/**
	 * Sets the column data extraction strategy. Columns are processed in order.
	 * 'skip' is a special strategy that skips over a column.
	 * 
	 * @param columns LinkedHashMap of column names to convertor names
	 */
	void columns(Map columns) { columnMap = columns }
	
	/**
	 * Adds or replaces a convertor.
	 * 
	 * @param name of the convertor
	 * @param convertor Closure that extracts data from a Cell
	 */
	void convertor(String name, Closure convertor) { convertors[name] = convertor }
	
	private List<Map> data(Workbook workbook) {
		List data = []
		Sheet sheet = workbook.getSheetAt(0)
		int rows = sheet.physicalNumberOfRows - startRowIndex
		rows.times {
			Row row = sheet.getRow(it + startRowIndex)
			if (row) { data << rowData(row) }
		}
		data
	}
	
	private Map rowData(Row row) {
		Map data = [:]
		columnMap.eachWithIndex { column, name, index ->
			Cell cell = row.getCell(index)
			if (name != 'skip') {
				Closure convertor = convertors[name]
				if (convertor) { data[column] = convertor cell }
				else { 
					throw new IllegalArgumentException("$name is not a supported convertor for column $column")
				}
			}
		}
		data
	}
	
	private String cellAsString(Cell cell) { cell.toString() }
	
	private Boolean cellAsBoolean(Cell cell) { cell.booleanCellValue }
	
	private BigDecimal cellAsBigDecimal(Cell cell) { new BigDecimal(cell.toString()) }

	private Double cellAsDouble(Cell cell) { cell.numericCellValue}
	
	private Float cellAsFloat(Cell cell) { cellAsDouble(cell).floatValue() }
	
	private Long cellAsLong(Cell cell) {  cellAsBigDecimal(cell).longValue() }
	
	private Integer cellAsInteger(Cell cell) { cellAsBigDecimal(cell).intValue() }
	
	private Date cellAsDate(Cell cell) { cell.dateCellValue }
	
}
