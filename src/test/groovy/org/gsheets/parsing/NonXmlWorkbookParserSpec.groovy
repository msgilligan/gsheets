package org.gsheets.parsing

import org.gsheets.NonXmlWorkbookSupport
import org.gsheets.building.WorkbookBuilder

class NonXmlWorkbookParserSpec extends WorkbookParserSpec {

	protected newParser() {
		new WorkbookParser(false)
	}
	
	protected newBuilder() {
		new WorkbookBuilder(false)
	}
	
	def 'can parse non Xml workbooks'() {
		expect:
		parser.support instanceof NonXmlWorkbookSupport
		builder.support instanceof NonXmlWorkbookSupport
	}

}
