package org.gsheets

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.After
import org.junit.Before

class WorkbookBuilderTests extends WorkbookBuilderTestCase {

	protected WorkbookBuilderSupport newBuilder() {
		builder = new WorkbookBuilder()
	}
	
	protected Class workbookType() {
		HSSFWorkbook
	}
	
}
