package org.gsheets

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Workbook

class NonXmlWorkbookSupport implements WorkbookSupport {
	
	Class<? extends Workbook> workbookType() { HSSFWorkbook }
	
}
