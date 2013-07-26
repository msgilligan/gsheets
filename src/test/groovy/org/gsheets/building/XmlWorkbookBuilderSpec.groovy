package org.gsheets.building

class XmlWorkbookBuilderSpec extends WorkbookBuilderSpec {

	protected newBuilder() { new WorkbookBuilder(true) }
	
	static void main(String[] args) {
		String filename = args.length ? args[0] : 'demo.xlsx'
		file  filename, new XmlWorkbookBuilderSpec().newBuilder()
	}
	
}
