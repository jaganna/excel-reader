package sandbox.groovy.excel

import groovy.transform.Canonical
import groovy.transform.ToString
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFWorkbook

class ExcelReader {

    public static final String FILE_NAME = './foo.xlsx'

    static void main(String[] args) {

        def positionParser = new PositionParser(FILE_NAME)
        positionParser.CONFIGURATIONS.keySet().each {
            println("Reading $it")
            try {
                def positions = positionParser.positionsFor(it)
                println(positions)
            } catch (any) {
                println("ERROR: " + any)
            }
            println()
        }
    }
}


@Canonical
@ToString(excludes = "complete")
class Position {
    final String usiPrefix
    final String usiValue
    final String party2

    Position(String usiPrefix, String usiValue, String party2) {
        this.usiPrefix = usiPrefix
        this.usiValue = usiValue
        this.party2 = party2
    }

    boolean isComplete() {
        notEmpty(usiValue) && notEmpty(usiPrefix) && notEmpty(party2)
    }

    private static boolean notEmpty(String value) {
        value != null && !value.isEmpty()
    }
}


interface RowToPosition {
    Position map(XSSFRow row)
}

@Canonical
class SheetConfiguration {
    String sheetName
    int startOffset
    int endOffset
    RowToPosition rowToPositionMapper
}

class PositionParser {

    static CONFIGURATIONS = [
            'FX'     : new SheetConfiguration(sheetName: 'Fx', startOffset: 3, endOffset: 6, rowToPositionMapper: new FxRowToPosition()),
            'EQUITY' : new SheetConfiguration(sheetName: 'Equity', startOffset: 3, endOffset: 4, rowToPositionMapper: new FxRowToPosition()),
            'RATES'  : new SheetConfiguration(sheetName: 'Rates', startOffset: 3, endOffset: 4, rowToPositionMapper: new RatesRowToPosition()),
            'CREDITS': new SheetConfiguration(sheetName: 'Creadits', startOffset: 3, endOffset: 4, rowToPositionMapper: new FxRowToPosition())
    ]

    private XSSFWorkbook workbook

    PositionParser(String fileName) {
        workbook = new XSSFWorkbook(new File(fileName).newInputStream())
    }

    List<Position> positionsFor(String assetClass) {
        List<Position> result = []

        def sheetConfiguration = CONFIGURATIONS.get(assetClass)
        def sheet = workbook.getSheet(sheetConfiguration.sheetName)

        def startRow = sheet.firstRowNum + sheetConfiguration.startOffset
        def endRow = sheet.lastRowNum - sheetConfiguration.endOffset
        startRow.upto(endRow) {
            def row = sheet.getRow(it)
            try {
                Position position = sheetConfiguration.rowToPositionMapper.map(row)
                if (position.isComplete())
                    result.add(position)
                else
                    println "error for row $it: " + position
            } catch (any) {
                println(any)
            }
        }

        result
    }
}


class FxRowToPosition implements RowToPosition {

    @Override
    Position map(XSSFRow row) {
        def prefix = row.getCell(1).stringCellValue
        def usiValue = row.getCell(2).stringCellValue
        def party2 = row.getCell(3).stringCellValue

        new Position(prefix, usiValue, party2)
    }
}

class RatesRowToPosition implements RowToPosition {

    @Override
    Position map(XSSFRow row) {
        def prefix = row.getCell(1).stringCellValue
        def usiValue = row.getCell(2).stringCellValue
        def party2 = row.getCell(3).stringCellValue

        new Position(prefix, usiValue, party2)
    }
}
