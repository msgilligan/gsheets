package org.gsheets.parsing

import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.gsheets.building.WorkbookBuilder

class XmlWorkbookParserSpec extends WorkbookParserSpec {

	protected WorkbookParser newParser(Workbook workbook) { new WorkbookParser(workbook) }
	
	protected WorkbookBuilder newBuilder() { new WorkbookBuilder(true) }
	
	// Note: run the corresponding builder spec as an app first to create the file 
	static void main(String[] args) {
		FileInputStream ins = new FileInputStream('demo.xlsx')
		Workbook workbook = new XSSFWorkbook(ins)
		WorkbookParser parser = new WorkbookParser(workbook)
		println parser.grid {
			headerRows 1
			columns name: 'int', date: 'date', count: 'int', value: 'decimal', active: 'boolean'
		}
		ins.close()
	}
	
}
