package org.gsheets

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Workbook

class WorkbookBuilder extends WorkbookBuilderSupport {
	
	protected Class workbookType() {
		HSSFWorkbook
	}
}
