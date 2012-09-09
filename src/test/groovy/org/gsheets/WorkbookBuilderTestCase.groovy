package org.gsheets

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.junit.Test

abstract class WorkbookBuilderTestCase {

	WorkbookBuilderSupport builder
	
	GroovyTestCase tc = new GroovyTestCase()

	Workbook getWb() {
		builder.wb
	}

	Sheet getSht() {
		builder.currentSheet
	}

	@Test
	void empty_workbook() {
		builder.workbook {
		}
		assert wb.numberOfSheets == 0
	}

	@Test
	void workbook_with_sheets() {
		builder.workbook {
			sheet('sheet1') {
			}
			sheet('sheet2') {
			}
		}

		wb.with {
			assert getSheetAt(0).sheetName == 'sheet1'
			assert getSheetIndex('sheet2') == 1
			assert numberOfSheets == 2
		}
	}
	
	@Test 
	void workbook_with_emptyRow() {
		builder.workbook {
			sheet('sheet') {
				emptyRow()
			}
		}
		assert builder.currentRow
		assert builder.currentRowNum == 1
	}

	@Test
	void workbook_with_header() {
		def headers = ['String', 'Boolean', 'Date', 'Object', 'Integer', 'Double', 'BigDecimal']

		builder.workbook {
			sheet('sheet') {
				header(headers)
			}
		}
		assertHeaderRow headers
	}
	
	@Test
	void workbook_header_must_be_firstRow() {
		def headers = ['String', 'Boolean', 'Date', 'Object', 'Integer', 'Double', 'BigDecimal']

		assert 'header must be first row' == tc.shouldFail(IllegalStateException) {
			builder.workbook {
				sheet('sheet') {
					emptyRow()
					header(headers)
				}
			}
		}
	}
		
	protected assert_stringCellValue(row, cellNum, value) {
		Cell cell = row.getCell(cellNum)
		assert cell.cellType == Cell.CELL_TYPE_STRING
		assert cell.stringCellValue == value.toString()
	}

	protected assert_dateCellValue(row, cellNum, Date value) {
		Cell cell = row.getCell(cellNum)
		assert cell.cellType == Cell.CELL_TYPE_NUMERIC
		assert cell.dateCellValue == value
	}

	protected assert_numericCellValue(row, cellNum, value) {
		Cell cell = row.getCell(cellNum)
		assert cell.cellType == Cell.CELL_TYPE_NUMERIC
		assert cell.numericCellValue == value
	}

	protected assertHeaderRow(headers) {
		Row header = sht.getRow(0)
		headers.eachWithIndex { h, i ->
			Cell cell = header.getCell(i)
			assert cell.cellType == Cell.CELL_TYPE_STRING
			assert cell.stringCellValue == h
		}
		assert builder.currentRowNum == 1
	}
}