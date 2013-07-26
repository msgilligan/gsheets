package org.gsheets.parsing

class Grid {

	int rows
	
	int startRow
	
	Map columns = [:]
	
	void header(int rows) { startRow = rows }
}
