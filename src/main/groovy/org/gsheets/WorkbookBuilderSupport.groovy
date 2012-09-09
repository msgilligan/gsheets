package org.gsheets

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook

/**
 * Base class for Workbook builders
 * 
 * @author Ken Krebs 
 */
abstract class WorkbookBuilderSupport {
	
	Workbook wb
	Sheet currentSheet
	int currentRowNum
	Row currentRow
	
	abstract protected Class workbookType()

	WorkbookBuilderSupport() {
		wb = workbookType().newInstance()
	}
	
	/**
	 * Builds a new workbook.
	 *
	 * @param the closure holds nested method calls
	 * 
	 * @return the created Workbook
	 */
	Workbook workbook(Closure closure) {
		assert closure

		closure.delegate = this
		closure.call()
		wb
	}

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
	
	Row emptyRow() {
		assert currentSheet
		
		currentRow = currentSheet.createRow(currentRowNum++)
	}

	Row header(List<String> names)  {
		assert currentSheet
		assert names
		if(currentRow != null) {
			throw new IllegalStateException('header must be first row')
		}

		currentRow = currentSheet.createRow(0)
		currentRowNum++
		names.eachWithIndex { String value, col ->
			Cell cell = currentRow.createCell(col)
			cell.setCellValue(value)
		}
		currentRow
	}


}
