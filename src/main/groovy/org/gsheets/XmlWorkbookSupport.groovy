package org.gsheets

import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook

class XmlWorkbookSupport implements WorkbookSupport {

	Class<? extends Workbook> workbookType() { XSSFWorkbook }
	
}
