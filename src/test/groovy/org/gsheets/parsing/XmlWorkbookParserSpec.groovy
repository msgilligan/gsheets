package org.gsheets.parsing

import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.gsheets.building.WorkbookBuilder

class XmlWorkbookParserSpec extends WorkbookParserSpec {

	protected WorkbookParser newParser(Workbook workbook) { new WorkbookParser(workbook) }
	
	protected WorkbookBuilder newBuilder() { new WorkbookBuilder(true) }
	
	static void main(String[] args) {
		FileInputStream ins = new FileInputStream('demo.xlxs')
		Workbook workbook = new XSSFWorkbook(ins)
		WorkbookParser parser = new WorkbookParser(workbook)
		println parser.grid {
			headerRows 1
			columns name: String, date: Date, count: Integer, value: BigDecimal, active: Boolean
		}
		ins.close()
	}
	
}
