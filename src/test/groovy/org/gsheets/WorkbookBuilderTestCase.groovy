package org.gsheets

import static org.junit.Assert.*

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Workbook
import org.junit.After
import org.junit.Before
import org.junit.Test

abstract class WorkbookBuilderTestCase {

	WorkbookBuilderSupport builder
	
	GroovyTestCase tc = new GroovyTestCase()
	
	def headers = ['a', 'b', 'c']
	
	abstract protected WorkbookBuilderSupport newBuilder()
	
	@Before()
	void setUp() {
		builder = newBuilder()
	}
	
	@After()
	void tearDown() {
		assert builder.wb.class == builder.workbookType()
	}
	
	@Test
	void empty_workbook() {
		builder.workbook {
		}
		assert builder.wb.numberOfSheets == 0
	}

	@Test
	void workbook_with_sheets() {
		builder.workbook {
			sheet('sheet1') {
			}
			sheet('sheet2') {
			}
		}

		builder.wb.with {
			assert getSheetAt(0).sheetName == 'sheet1'
			assert getSheetIndex('sheet2') == 1
			assert numberOfSheets == 2
		}
	}
	
	@Test
	void workbook_with_rows() {
		int v = 13
		builder.workbook {
			sheet('sheet') {
				row()
				row(headers)
				row('a', 'b', 'c')
				row("$v")
				row(true)
				row(12.3F)
				row(12.4D)
				row(12.34)
				row(v)
				row(v as Long)
				row(v as Short)
				row('x', v)
			}
		}
		
		assert cell0(0) == null
		assert_string_row(1, headers)
		assert_string_row(2, headers)
		assert cell0(3).stringCellValue == '13'
		assert cell0(4).booleanCellValue
		assertEquals cell0(5).numericCellValue, 12.3, 0.00001
		assertEquals cell0(6).numericCellValue, 12.4, 0.00001
		assertEquals cell0(7).numericCellValue, 12.34, 0.00001
		assert cell0(8).numericCellValue == 13
		assert cell0(9).numericCellValue == 13
		assert cell0(10).numericCellValue == 13
		assert cell(11, 0).stringCellValue == 'x' 
		assert cell(11, 1).numericCellValue == 13 
		assert builder.currentRow.rowNum == 11
		assert builder.nextRowNum == 12
	}
		
	protected assert_string_row(rowNum, values) {
		Row row = builder.currentSheet.getRow(rowNum)
		values.eachWithIndex { v, i ->
			Cell cell = row.getCell(i)
			assert cell.cellType == Cell.CELL_TYPE_STRING
			assert cell.stringCellValue == v
		}
	}
	
	protected Cell cell(rowNum, cellNum) {
		cell(builder.currentSheet.getRow(rowNum), cellNum)
	}
	
	protected Cell cell(Row row, cellNum) {
		row.getCell cellNum
	}

	protected Cell cell0(rowNum) {
		cell rowNum, 0
	}

	protected Cell cell(cellNum) {
		cell builder.currentRow, cellNum
	}

}