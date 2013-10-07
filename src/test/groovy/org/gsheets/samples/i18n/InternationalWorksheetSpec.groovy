package org.gsheets.samples.i18n

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Workbook
import org.gsheets.parsing.WorkbookParser
import spock.lang.Specification


class InternationalWorksheetSpec extends Specification {

    def 'can parse i18n sample workbook'() {
        given:
        FileInputStream ins = new FileInputStream('samples/i18n/i18n-master.xls')
        Workbook workbook = new HSSFWorkbook(ins)
        WorkbookParser parser = new WorkbookParser(workbook)

        when:
        List data = parser.grid {
            startRowIndex = 1
            columns key: 'string', en: 'string', es: 'string'
        }

        then:
        with(data[0]) {
            key == 'say.hello'
            en == "Hello"
            es == "Hola"
        }

        with(data[1]) {
            key == 'say.goodbye'
            en == "Goodbye"
            es == "Adios"
        }
    }
}