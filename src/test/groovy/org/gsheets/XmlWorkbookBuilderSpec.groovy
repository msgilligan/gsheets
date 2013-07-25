package org.gsheets

class XmlWorkbookBuilderSpec extends WorkbookBuilderSpec {

	protected newBuilder() {
		new WorkbookBuilder(true)
	}
	
	static void main(String[] args) {
		file 'demo.xlsx', new XmlWorkbookBuilderSpec().newBuilder()
	}
	
}
