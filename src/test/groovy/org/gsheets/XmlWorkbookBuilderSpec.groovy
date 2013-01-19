package org.gsheets

class XmlWorkbookBuilderSpec extends WorkbookBuilderSupportSpec {

	protected WorkbookBuilderSupport newBuilder() {
		builder = new XmlWorkbookBuilder()
	}
	
	static void main(String[] args) {
		file 'demo.xlsx', new XmlWorkbookBuilder()
	}
	
}
