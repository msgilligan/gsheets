package org.gsheets

import org.apache.poi.ss.usermodel.Workbook

class WorkbookBuilderSpec extends WorkbookBuilderSupportSpec {

	protected WorkbookBuilderSupport newBuilder() {
		builder = new WorkbookBuilder()
	}
	
	static void main(String[] args) {
		file 'demo.xls', new WorkbookBuilder()
	}
	
}
