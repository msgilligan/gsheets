package org.gsheets

import static org.junit.Assert.*

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
		assert wb.class == XSSFWorkbook
	}

}
