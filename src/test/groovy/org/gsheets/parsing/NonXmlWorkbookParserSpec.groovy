package org.gsheets.parsing

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Workbook
import org.gsheets.building.WorkbookBuilder

class NonXmlWorkbookParserSpec extends WorkbookParserSpec {

	protected WorkbookParser newParser(Workbook workbook) { new WorkbookParser(workbook) }
	
	protected WorkbookBuilder newBuilder() { new WorkbookBuilder(false) }
	
	static void main(String[] args) {
		FileInputStream ins = new FileInputStream('demo.xls')
		Workbook workbook = new HSSFWorkbook(ins)
		WorkbookParser parser = new WorkbookParser(workbook)
		println parser.grid {
			headerRows 1
			columns name: String, date: Date, count: Integer, value: BigDecimal, active: Boolean
		}
		ins.close()
	}
}
