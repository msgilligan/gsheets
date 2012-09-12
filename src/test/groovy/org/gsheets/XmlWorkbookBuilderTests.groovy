package org.gsheets

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.junit.After
import org.junit.Before

class XmlWorkbookBuilderTests extends WorkbookBuilderTestCase {

	protected WorkbookBuilderSupport newBuilder() {
		builder = new XmlWorkbookBuilder()
	}
	
	protected Class workbookType() {
		XSSFWorkbook
	}
	
}
