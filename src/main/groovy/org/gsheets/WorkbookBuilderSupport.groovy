package org.gsheets

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook

/**
 * Base class for Workbook builders.
 * 
 * @author Ken Krebs 
 */
abstract class WorkbookBuilderSupport {

	Workbook wb
	Sheet currentSheet
	int nextRowNum
	Row currentRow

	abstract protected Class workbookType()

	protected WorkbookBuilderSupport() {
		wb = workbookType().newInstance()
	}

	/**
	 * Provides the root of a Workbook DSL.
	 *
	 * @param the closure supports nested method calls
	 * 
	 * @return the Workbook
	 */
	Workbook workbook(Closure closure) {
		assert closure

		closure.delegate = this
		closure.call()
		wb
	}

	/**
	 * Builds a new Sheet.
	 *
	 * @param the closure supports nested method calls
	 * 
	 * @return the created Sheet
	 */
	Sheet sheet(String name, Closure closure) {
		assert wb
		assert name
		assert closure

		currentSheet = wb.createSheet(name)
		assert currentSheet
		closure.delegate = currentSheet
		closure.call()
		currentSheet
	}

	Row row() {
		currentRow = currentSheet.createRow(nextRowNum++)		
	}
	
	Row row(Object value) {
		row([value])
	}
	
	Row row(... values) {
		assert currentSheet

		row()
		if (values) {
			values.eachWithIndex { value, index ->
				cell value, index
			}
		}
		currentRow
	}
	
	Row row(List values) {
		row(values.toArray())
	}
	
	Cell cell(String value, int index) {
		cell value, index, Cell.CELL_TYPE_STRING 
	}
	
	Cell cell(Boolean value, int index) {
		cell value, index, Cell.CELL_TYPE_BOOLEAN
	}
	
	Cell cell(Number value, int index) {
		cell value, index, Cell.CELL_TYPE_NUMERIC
	}
	
	Cell cell(Date date, int index) {
		cell date, index, Cell.CELL_TYPE_NUMERIC
	}
	
	Cell cell(Formula formula, int index) {
		cell formula.text, index, Cell.CELL_TYPE_FORMULA
	}
	
	Cell cell(value, int index) {
		cell value.toString(), index, Cell.CELL_TYPE_STRING
	}
	
	protected Cell cell(value, int index, int cellType) {
		Cell cell = currentRow.createCell(index)
		cell.cellType = cellType
		cell.setCellValue(value)
		cell
	}
}

