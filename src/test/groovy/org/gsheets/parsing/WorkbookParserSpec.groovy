package org.gsheets.parsing

import org.apache.poi.ss.usermodel.Workbook
import org.gsheets.building.WorkbookBuilder

import spock.lang.Specification

abstract class WorkbookParserSpec extends Specification {

	WorkbookParser parser
	
	WorkbookBuilder builder
	
	abstract protected newParser()
	
	abstract protected newBuilder()
	
	def setup() {
		parser = newParser()
		builder = newBuilder()
	}
	
	def 'has a startRowIndex'() {
		expect: parser.startRowIndex == 0
		
		when: parser.header(2)
		
		then: parser.startRowIndex == 2
	}
	
	def 'has column Map of names to typeConvertors'() {
		when:
		parser.columnMap = [x: String]
		
		then:
		parser.columnMap.x == String
	}

	def 'can parse a grid without a header of simple types originating from row 0 column 0 from the first worksheet'() {
		given:
		Date date1 = new Date()
		Date date2 = new Date(date1.time + 1000 * 60 * 60 * 24)
		builder.workbook {
			sheet('a sheet') {
				row('a', 12.34, 13, true, 3.14159D, 2.345F, date1, 123456789L)
				row('b', -98.62, -69, false, 2.71821828D, 13.01F, date2, 999999999999999L)
			}
		}
		
		when: 'columns are Maps of String to type'
		List data = parser.grid(builder.wb) {
			columns abbreviation: String, cost: BigDecimal, num: Integer, status: Boolean, irr: Double, flt: Float, date: Date, lng: Long
		}
		
		then:
		parser.workbook == builder.wb
		data == [
			[abbreviation: 'a', cost: 12.34, num: 13, status: true, irr: 3.14159D, flt: 2.345F, date: date1, lng: 123456789L],
			[abbreviation: 'b', cost: -98.62, num: -69, status: false, irr: 2.71821828D, flt: 13.01F, date: date2, lng: 999999999999999L],
		]
		
	}

	def 'can parse a grid with ignored header rows of simple types originating from row 0 column 0 from the first worksheet'() {
		given:
		Date date1 = new Date()
		Date date2 = new Date(date1.time + 1000 * 60 * 60 * 24)
		builder.workbook {
			sheet('a sheet') {
				row('Abbreviation', 'Cost', '')
				row('ab', 'x', 'y')
				row('a', 12.34, 13, true, 3.14159D, 2.345F, date1, 123456789L)
				row('b', -98.62, -69, false, 2.71821828D, 13.01F, date2, 999999999999999L)
			}
		}
		
		when: 'columns are Maps of String to type'
		List data = parser.grid(builder.wb) {
			header 2
			columns abbreviation: String, cost: BigDecimal, num: Integer, status: Boolean, irr: Double, flt: Float, date: Date, lng: Long
		}
		
		then:
		parser.workbook == builder.wb
		data == [
			[abbreviation: 'a', cost: 12.34, num: 13, status: true, irr: 3.14159D, flt: 2.345F, date: date1, lng: 123456789L],
			[abbreviation: 'b', cost: -98.62, num: -69, status: false, irr: 2.71821828D, flt: 13.01F, date: date2, lng: 999999999999999L],
		]
	}

	def 'can partially parse a grid with not enough rows'() {
		
	}
}
