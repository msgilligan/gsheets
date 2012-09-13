package org.gsheets

class XmlWorkbookBuilderSpec extends WorkbookBuilderBaseSpec {

	protected WorkbookBuilderSupport newBuilder() {
		builder = new XmlWorkbookBuilder()
	}
	
	static void main(String[] args) {
		file 'demo.xlsx', new XmlWorkbookBuilder()
	}
	
}
