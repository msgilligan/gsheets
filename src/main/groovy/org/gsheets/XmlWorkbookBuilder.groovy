package org.gsheets

import org.apache.poi.xssf.usermodel.XSSFWorkbook

class XmlWorkbookBuilder extends WorkbookBuilderSupport {

	protected Class workbookType() {
		XSSFWorkbook
	}
}
