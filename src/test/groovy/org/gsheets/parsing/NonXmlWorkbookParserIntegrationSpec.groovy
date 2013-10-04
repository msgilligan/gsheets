package org.gsheets.parsing

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Workbook

import spock.lang.Specification

class NonXmlWorkbookParserIntegrationSpec extends Specification {

	def 'can parse a non xml workbook'() {
		given:
        String tzShort = TimeZone.default.getDisplayName(true, TimeZone.SHORT)
		FileInputStream ins = new FileInputStream('parsing_demo.xls')
		Workbook workbook = new HSSFWorkbook(ins)
		WorkbookParser parser = new WorkbookParser(workbook)
		
		when:
		List data = parser.grid {
			startRowIndex = 1
			columns name: 'string', date: 'date', count: 'int', value: 'decimal', active: 'boolean'
		}
		
		then:
		with(data[0]) {
			name == 'a'
			date.toString() == "Wed Sep 12 00:00:00 ${tzShort} 2012"
			count == 69
			value == 12.34
			active
		}
		
		with(data[1]) {
			name == 'b'
			date.toString() == "Thu Sep 13 00:00:00 ${tzShort} 2012"
			count == 666
			value == 43.21
			!active
		}
	}
}
