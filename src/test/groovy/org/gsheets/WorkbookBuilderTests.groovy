package org.gsheets

import static org.junit.Assert.*

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.junit.After
import org.junit.Before

class WorkbookBuilderTests extends WorkbookBuilderTestCase {

	@Before()
	void setup() {
		builder = new WorkbookBuilder()
	}
	
	@After
	void teardown() {
		assert wb.class == HSSFWorkbook
	}

}
