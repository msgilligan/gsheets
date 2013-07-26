package org.gsheets.parsing

import org.gsheets.XmlWorkbookSupport
import org.gsheets.building.WorkbookBuilder

class XmlWorkbookParserSpec extends WorkbookParserSpec {

	protected newParser() {
		new WorkbookParser(true)
	}
	
	protected newBuilder() {
		new WorkbookBuilder(true)
	}
	
	def 'can parse Xml workbooks'() {
		expect:
		parser.support instanceof XmlWorkbookSupport
		builder.support instanceof XmlWorkbookSupport
	}

}
