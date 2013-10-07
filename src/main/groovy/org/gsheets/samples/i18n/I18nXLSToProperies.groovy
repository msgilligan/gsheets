package org.gsheets.samples.i18n

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Workbook
import org.gsheets.parsing.WorkbookParser

class I18nXLSToProperties {

    File input
    File outDir
    String baseName
    Map i18nCols

    static List languages = ["en", "es"]  // Columns starting with 'B' (or '1')

    static void main(String[] args) {
        def splitter = new I18nXLSToProperties("samples/i18n/i18n-master.xls", "build/sample-output", "Strings")
        splitter.split()
    }

    I18nXLSToProperties(String xlsFile, String outputDir, String baseName) {
        input = new File(xlsFile)
        outDir = new File(outputDir)
        this.baseName = baseName

        i18nCols = [key: 'string']
        languages.each{ i18nCols.put(it, 'string') }
    }

    void split() {
        FileInputStream ins = new FileInputStream(input)
        Workbook workbook = new HSSFWorkbook(ins)
        WorkbookParser parser = new WorkbookParser(workbook)
        List data = parser.grid {
            startRowIndex = 1
            columns i18nCols
        }
        ins.close()

        languages.each { String lang ->
            Properties props = new Properties()
            File propsFile = new File(outDir, baseName + "_" + lang + ".properties")

            data.each { row ->
                props.setProperty(row.key, row[lang])
            }
            props.store(propsFile.newWriter(), null)
        }
    }

}
