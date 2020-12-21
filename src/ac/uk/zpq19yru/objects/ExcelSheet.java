package ac.uk.zpq19yru.objects;

/*
    
    Created By:     Callum Johnson
    Created In:     Dec/2020
    Project Name:   Payroll Collator
    Package Name:   ac.uk.zpq19yru.objects
    Class Purpose:  Excel Sheet Bridge, enables specific Data Setting/Pulling Methodology.
    
*/

import ac.uk.zpq19yru.exceptions.CellNotFoundException;
import ac.uk.zpq19yru.exceptions.CellTypeInvalidException;
import ac.uk.zpq19yru.exceptions.SheetNotValidException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.LinkedList;
import java.util.Optional;

public class ExcelSheet {

    private final Sheet sheet;
    private LinkedList<Cell> sheetData = new LinkedList<>();
    private final int maxRows;
    private final int loadedCells;
    private final LinkedList<Integer> columnData = new LinkedList<>();

    /**
     * Constructor to initialise an ExcelSheet.
     *
     * @param sheet - Worksheet to be Bridged.
     * @throws SheetNotValidException - throws when:
     *                                               Sheet is Null.
     *                                               Sheet is Empty.
     */
    public ExcelSheet(Sheet sheet, boolean fresh) throws SheetNotValidException {
        if (sheet == null) {
            throw new SheetNotValidException("Sheet is Null.");
        }
        if (sheet.getPhysicalNumberOfRows() <= 0 && !fresh) {
            throw new SheetNotValidException("Sheet is empty.");
        }
        this.sheet = sheet;
        for (int row = 0; row < sheet.getPhysicalNumberOfRows(); row++) {
            Row rowObject = sheet.getRow(row);
            if (rowObject == null) {
                rowObject = sheet.createRow(row);
            }
            for (int col = 0; col < rowObject.getPhysicalNumberOfCells(); col++) {
                sheetData.add(rowObject.getCell(col));
                if (columnData.size() == row) {
                    columnData.add(rowObject.getPhysicalNumberOfCells());
                }
            }
        }
        this.maxRows = sheet.getPhysicalNumberOfRows();
        loadedCells = sheetData.size();
    }

    /**
     * Method to return the Worksheet Name.
     *
     * @return - Bridged Worksheet Name.
     */
    public String getName() {
        return sheet.getSheetName();
    }

    /**
     * Method to return the maximum amount of Rows the document has.
     *
     * @return - max rows.
     */
    public int getMaxRows() {
        return maxRows;
    }

    /**
     * Method to return the maximum amount of Columns the document has at a specific row.
     *
     * @param row - Row index.
     * @return - Max Columns.
     */
    public int getMaxColumns(int row) {
        return columnData.get(row);
    }

    /**
     * Method to return the amount of Cells loaded into storage from the linked ExcelSheet.
     *
     * @return - Amount of Cells successfully loaded.
     */
    public int getLoadedCells() {
        return loadedCells;
    }

    /**
     * Method to return a Cell from the Workbook.
     *
     * @param row - Row of the Cell to be found.
     * @param column - Column of the Cell to be found.
     * @return - Cell Object.
     * @throws CellNotFoundException - Thrown if the Cell isn't found.
     */
    public Cell getCell(int row, int column) throws CellNotFoundException {
        Optional<Cell> cellOptional = sheetData.stream().filter(
                cell -> cell.getRowIndex() == row && cell.getColumnIndex() == column
        ).findFirst();
        if (cellOptional.isPresent()) {
            return cellOptional.get();
        } else {
            throw new CellNotFoundException("The cell at X'" + column + "' and Y'" + row + "' could not be found!");
        }
    }

    /**
     * Method to set a Cell's value.
     *
     * @param row - Row of the Cell to be found.
     * @param column - Column of the Cell to be found.
     * @param value - Value to be Set.
     * @param type - CellType you desire to create.
     * @param make - If the Cell isn't found, force-create the cell?
     * @throws CellNotFoundException - If the Cell isn't Found & Make is False.
     * @throws CellTypeInvalidException - If the CellType isn't recognised.
     */
    public void setCell(
            int row, int column, Object value, CellType type, boolean make
    ) throws CellNotFoundException, CellTypeInvalidException {
        Cell cell;
        try {
            cell = getCell(row, column);
        } catch (CellNotFoundException e) {
            if (!make) {
                throw e;
            }
            Row temp = sheet.getRow(row);
            if (temp == null) {
                temp = sheet.createRow(row);
            }
            cell = temp.getCell(column, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            sheetData.add(cell);
        }
        switch (type) {
            case NUMERIC:
                cell.setCellValue((Double) value);
                break;
            case STRING:
                cell.setCellValue((String) value);
                break;
            case FORMULA:
                cell.setCellFormula((String) value);
                break;
            case BOOLEAN:
                cell.setCellValue((Boolean) value);
                break;
            default:
                throw new CellTypeInvalidException("The CellType '" + type.name() + "' isn't supported.");
        }
    }

}
