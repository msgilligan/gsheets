package org.gsheets

import static org.hamcrest.Matchers.*
import static spock.util.matcher.HamcrestSupport.*

import java.text.SimpleDateFormat

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook

import spock.lang.Specification

abstract class WorkbookBuilderSupportSpec extends Specification {
	
	WorkbookBuilderSupport builder
	
	def headers = ['a', 'b', 'c']
	
	abstract protected WorkbookBuilderSupport newBuilder()
	
	def setup() {
		builder = newBuilder()
	}

	def cleanup() {
		assert builder.wb.class == builder.workbookType()
	}
	
	def 'can build an empty workbook'() {
		when:
		builder.workbook {
		}

		then:
		wb.numberOfSheets == 0
	}
	
	def 'can build a workbook with multiple sheets'() {
		when:
		builder.workbook {
			sheet('sheet1') {
			}
			sheet('sheet2') {
			}
		}
		
		then:
		wb.getSheetAt(0).sheetName == 'sheet1'
		wb.getSheetAt(1).sheetName == 'sheet2'
	}
	
	def 'can build 1 empty row'() {
		when:
		builder.workbook {
			sheet('a') {
				row()
			}
		}
		
		then:
		cell0(0) == null
		sheet.lastRowNum == 0
	}
	
	def 'can build multiple rows'() {
		when:
		builder.workbook {
			sheet('b') {
				row()
				row()
			}
		}
		
		then:
		cell0(0) == null
		cell0(1) == null
		sheet.lastRowNum == 1
	}
	
	def 'can build a String cell'() {
		when:
		builder.workbook {
			sheet('c') {
				row('s')
			}
		}
		
		then:
		cell0(0).stringCellValue == 's'
	}
	
	def 'can build a Boolean cell'() {
		when:
		builder.workbook {
			sheet('d') {
				row(true)
				row(false)
			}
		}
		
		then:
		cell0(0).booleanCellValue
		!cell0(1).booleanCellValue
	}
	
	def 'can build a Double cell'() {
		when:
		builder.workbook {
			sheet('e') {
				row(12.345D)
				row(13)
				row(14.18)
				row(23.45F)
				row(123456789L)
				row(123 as Short)
			}
		}
		
		then:
		cell0(0).numericCellValue == 12.345D
		cell0(1).numericCellValue == 13D
		cell0(2).numericCellValue == 14.18D
		that cell0(3).numericCellValue, closeTo(23.45D, 0.000001D)
		cell0(4).numericCellValue == 123456789
		cell0(5).numericCellValue == 123
	}
	
	def 'can build a Date cell'() {
		given:
		Date date = new Date()
		
		when:
		builder.workbook {
			sheet('f') {
				row(date)
			}
		}
		
		then:
		cell0(0).dateCellValue == date
	}
	
	def 'can build an object cell as a String'() {
		given:
		SomeClass someObject = new SomeClass()
		
		when:
		builder.workbook {
			sheet('g') {
				row(someObject)
			}
		}
		
		then:
		cell0(0).stringCellValue == someObject.toString()

	}
	
	def 'can build a Formula cell'() {
		given:
		Formula formula = new Formula('A1*B1')
		
		when:
		builder.workbook {
			sheet('h') {
				row(13, 3, formula)
			}
		}
		
		then:
		cell(0, 0).numericCellValue == 13
		cell(0, 1).numericCellValue == 3
		cell(0, 2).stringCellValue == 'A1*B1'
		cell(0, 2).cellType == Cell.CELL_TYPE_FORMULA
	}
	
	def 'can build multiple cells in a row'() {
		when:
		builder.workbook {
			sheet('i') {
				row(13, 'x', true)
			}
		}
		
		then:
		cell(0, 0).numericCellValue == 13
		cell(0, 1).stringCellValue == 'x'
		cell(0, 2).booleanCellValue
	}
	
	def 'cannot build a row outside a sheet'() {
		when:
		builder.workbook {
			row()
		}
		
		then:
		thrown(NullPointerException)
	}
	
	static class SomeClass {
		String toString() {
			'someObject'
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

	protected Workbook getWb() {
		builder.wb
	}
	
	protected Sheet getSheet() {
		builder.currentSheet
	}
	
	static demospec = {
		workbook {
			def fmt = new SimpleDateFormat('yyyy-MM-dd', Locale.default)
			sheet('sheet 1') {
				row('Name', 'Date', 'Count', 'Value', 'Active')
				row('a', fmt.parse('2012-09-12'), 69, 12.34, true)
				row('b', fmt.parse('2012-09-13'), 666, 43.21, false)
			}
		}
	}
	
	static file(name, builder) {
		Workbook workbook = builder.workbook demospec
		
		File file = new File(name)
		if (!file.exists()) {
			file.createNewFile()
		}
		def out = new FileOutputStream(file)
		workbook.write out
		out.close()
	}
}
