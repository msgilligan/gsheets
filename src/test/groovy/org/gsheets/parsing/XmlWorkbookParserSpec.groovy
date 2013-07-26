package org.gsheets.parsing

import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.gsheets.XmlWorkbookSupport
import org.gsheets.building.WorkbookBuilder

class XmlWorkbookParserSpec extends WorkbookParserSpec {

	protected newParser() {
		new WorkbookParser(true)
	}
	
	protected newBuilder() {
		new WorkbookBuilder(true)
	}
	
	def 'can parse Xml workbooks'() {
		expect:
		parser.support instanceof XmlWorkbookSupport
		builder.support instanceof XmlWorkbookSupport
	}
	
	static void main(String[] args) {
		FileInputStream ins = new FileInputStream('demo.xlsx')
		Workbook workbook = new XSSFWorkbook(ins)
		WorkbookParser parser = new WorkbookParser(true)
		println parser.grid(workbook) {
			header 1
			columns name: String, date: Date, count: Integer, value: BigDecimal, active: Boolean
		}
		ins.close()
	}
	
}
