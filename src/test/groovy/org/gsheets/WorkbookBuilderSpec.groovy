package org.gsheets

import java.text.SimpleDateFormat

import org.apache.poi.ss.formula.functions.T
import org.apache.poi.ss.usermodel.Workbook

class WorkbookBuilderSpec extends WorkbookBuilderBaseSpec {

	protected WorkbookBuilderSupport newBuilder() {
		builder = new WorkbookBuilder()
	}
	
	static void main(String[] args) {
		file 'demo.xls', new WorkbookBuilder()
	}
	
}
