package org.gsheets

import org.apache.poi.ss.usermodel.Workbook

class NonXmlWorkbookBuilderSpec extends WorkbookBuilderSpec {

	protected newBuilder() {
		new WorkbookBuilder(false)
	}
	
	static void main(String[] args) {
		file 'demo.xls', new NonXmlWorkbookSupport().newBuilder()
	}
	
}
