package org.gsheets.parsing

import spock.lang.Specification

class GridSpec extends Specification {
	
	Grid grid = new Grid()
	
	def 'has rows'() {
		when:
		grid.rows = 2
		
		then:
		grid.rows == 2
		grid.startRow == 0
	}

	def 'has a startRow'() {
		when:
		grid.startRow = 1
		
		then:
		grid.startRow == 1
	}

	def 'has a header'() {
		when:
		grid.header(2)
		
		then:
		grid.startRow == 2
	}

	def 'has column Map of names to types'() {
		when:
		grid.columns = [x: String]
		
		then:
		grid.columns.x == String
	}

}
