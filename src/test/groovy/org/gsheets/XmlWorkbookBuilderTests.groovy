package org.gsheets

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.junit.After
import org.junit.Before

class XmlWorkbookBuilderTests extends WorkbookBuilderTestCase {

	@Before()
	void setup() {
		builder = new XmlWorkbookBuilder()
	}
	
	@After
	void teardown() {
		assert builder.wb.class == XSSFWorkbook
	}

}
