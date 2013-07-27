package org.gsheets.parsing

import org.apache.poi.ss.usermodel.Workbook
import org.gsheets.building.WorkbookBuilder

import spock.lang.Specification
import spock.lang.Unroll

abstract class WorkbookParserSpec extends Specification {

	WorkbookParser parser
	
	WorkbookBuilder builder
	
	abstract protected WorkbookParser newParser(Workbook workbook)
	
	abstract protected WorkbookBuilder newBuilder()
	
	Date date1 = new Date()
	Date date2 = new Date(date1.time + 1000 * 60 * 60 * 24)
	List rowA = ['a', 'rubbish', 12.34, 13, true, 3.14159D, 2.345F, date1, 123456789L]
	List rowB = ['b', 'garbage', -98.62, -69, false, 2.71821828D, 13.01F, date2, 999999999999999L]
	Map mapA = [
		abbreviation: 'a', cost: 12.34, num: 13, status: true,
		irr: 3.14159D, flt: 2.345F, date: date1, lng: 123456789L
	]
	Map mapB = [
		abbreviation: 'b', cost: -98.62, num: -69, status: false,
		irr: 2.71821828D, flt: 13.01F, date: date2, lng: 999999999999999L
	]
	
	Closure buildIt = {
		sheet('a sheet') {
			row rowA
			row rowB
		}
	}
	
	Closure gridIt = {
		columns(
			abbreviation: 'string', disregard: 'skip', cost: 'decimal',
			num: 'int', status: 'boolean', irr: 'double', flt: 'float', date: 'date', lng: 'long'
		)
	}
	
	def setup() { builder = newBuilder() }
	
	def 'can parse a grid without header rows of simple types'() {
		given:
		builder.workbook buildIt
		parser = newParser(builder.wb)
		
		when:
		List data = parser.grid gridIt
		
		then:
		data == [mapA, mapB]
	}

	@Unroll
	def 'from parsed grid: #a plus #b equals #c'() {
		expect:
		a + b == c
		
		where:
		row << abcData()
		a = row.a
		b = row.b
		c = row.c
	}
	
	private List abcData() {
		setup()
		builder.workbook {
			sheet('a sheet') {
				row 2, 4, 6
				row 23, 46, 69
			}
		}
		
		newParser(builder.wb).grid { columns a: 'int', b: 'int', c: 'int' }
	}
	
	def 'can parse a grid with ignored header rows of simple types'() {
		given:
		builder.workbook {
			sheet('a sheet') {
				row('Abbreviation', 'Cost', '')
				row('ab', 'x', 'y')
				row rowA
				row rowB
			}
		}
		parser = newParser(builder.wb)
		
		when:
		List data = parser.grid {
			headerRows 2
			columns(
				abbreviation: 'string', disregard: 'skip', cost: 'decimal',
				num: 'int', status: 'boolean', irr: 'double', flt: 'float', date: 'date', lng: 'long'
			)
		}
		parser = newParser(builder.wb)
		
		then:
		data == [mapA, mapB]
	}
	
	def 'can NOT parse a grid with an unsupported or unknown conversion'() {
		given:
		builder.workbook { 
			sheet('a') {
				row 'whatever'
			}
		}
		parser = newParser(builder.wb)
		
		when: 
		parser.grid { columns y: 'map' }
		
		then:
		IllegalArgumentException x = thrown()
		x.message == 'map is not a supported convertor for column y'
	}
	
	def 'can parse using a custom convertor'() {
		given:
		builder.workbook {
			sheet('special') {
				row 'something'
				row 'else'
			}
		}
		parser = newParser(builder.wb)
		
		when:
		List data = parser.grid {
			convertor('toUpper') { it.toString().toUpperCase() }
			columns x: 'toUpper'
		}
		
		then:
		data == [[x: 'SOMETHING'], [x: 'ELSE']]
	}

}
