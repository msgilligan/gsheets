package org.gsheets;

import org.apache.poi.ss.usermodel.Workbook;

public interface WorkbookSupport {

	Class<? extends Workbook> workbookType();
	
}
