package org.gsheets.parsing

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Workbook
import org.gsheets.NonXmlWorkbookSupport
import org.gsheets.building.WorkbookBuilder

class NonXmlWorkbookParserSpec extends WorkbookParserSpec {

	protected newParser() {
		new WorkbookParser(false)
	}
	
	protected newBuilder() {
		new WorkbookBuilder(false)
	}
	
	def 'can parse non Xml workbooks'() {
		expect:
		parser.support instanceof NonXmlWorkbookSupport
		builder.support instanceof NonXmlWorkbookSupport
	}

	static void main(String[] args) {
		FileInputStream ins = new FileInputStream('demo.xls')
		Workbook workbook = new HSSFWorkbook(ins)
		WorkbookParser parser = new WorkbookParser(false)
		println parser.grid(workbook) {
			header 1
			columns name: String, date: Date, count: Integer, value: BigDecimal, active: Boolean
		}
		ins.close()
	}
}
