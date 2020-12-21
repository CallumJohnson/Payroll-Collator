package ac.uk.zpq19yru.objects;

/*
    
    Created By:     Callum Johnson
    Created In:     Dec/2020
    Project Name:   Payroll Collator
    Package Name:   ac.uk.zpq19yru.objects
    Class Purpose:  Excel Document Bridge, enables specific Data Pulling Methodology.
    
*/

import ac.uk.zpq19yru.Main;
import ac.uk.zpq19yru.exceptions.SheetNotValidException;
import ac.uk.zpq19yru.exceptions.WorkbookNotValidException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Iterator;
import java.util.LinkedList;

public class ExcelDocument implements Comparable<ExcelDocument> {

    private final Workbook workbook;
    private final LinkedList<ExcelSheet> sheets = new LinkedList<>();
    private final String workBookName;

    /**
     * Constructor to initialise an ExcelDocument.
     *
     * @param workbook - Workbook to be Bridged to an ExcelDocument.
     * @throws WorkbookNotValidException - Thrown when:
     *                                   Document is Null.
     *                                   Any sheet is Null.
     *                                   Document has no Pages.
     *                                   Documents Pages isn't equal to Found pages.
     */
    public ExcelDocument(Workbook workbook, String name) throws WorkbookNotValidException {
        if (workbook == null) {
            throw new WorkbookNotValidException("Workbook is Null.");
        }
        if (workbook.getNumberOfSheets() == 0 && !name.equals(Main.OUTPUT_FILE_NAME)) {
            throw new WorkbookNotValidException("Workbook has no sheets!");
        }
        this.workBookName = name;
        this.workbook = workbook;
        Iterator<Sheet> sheets = workbook.sheetIterator();
        while (sheets.hasNext()) {
            Sheet sheet = sheets.next();
            if (sheet == null) {
                throw new SheetNotValidException("Sheet is Null.");
            }
            ExcelSheet excelSheet = new ExcelSheet(sheet, false);
            System.out.println(
                    "Loaded '" + excelSheet.getLoadedCells()
                            + "' cells from the sheet '" + excelSheet.getName()
                            + "' from the workbook '" + name + "'!"
            );
            this.sheets.add(excelSheet);
        }
        if (workbook.getNumberOfSheets() != this.sheets.size()) {
            throw new WorkbookNotValidException("Workbook sheets could not be pulled from the Workbook!");
        }
    }

    /**
     * Get the workbook (HSSF or XSSF) linked to this ExcelDocument.
     *
     * @return - Workbook Object.
     */
    public Workbook getWorkbook() {
        return workbook;
    }

    public String getWorkBookName() {
        return workBookName;
    }

    /**
     * Method to return an ExcelSheet at an Index.
     *
     * @param index - Index of desired page.
     * @return - ExcelSheet bridge (Sheet -> ExcelSheet).
     * @throws IllegalArgumentException - If the Index is less than 0 or greater than the workbook's page count.
     */
    public ExcelSheet getSheet(int index) throws IllegalArgumentException {
        if (index < 0 || index > workbook.getNumberOfSheets()) {
            throw new IllegalArgumentException("Index provided is Greater than or Less than plausible Options.");
        }
        return sheets.get(index);
    }

    /**
     * Method to return an ExcelSheet by Name.
     *
     * @param name - Name of the desired page.
     * @return - ExcelSheet bridge (Sheet -> ExcelSheet).
     * @throws NullPointerException - If Name is Null or the page isn't found.
     */
    public ExcelSheet getSheet(String name) throws NullPointerException {
        if (name == null) {
            throw new NullPointerException("Workbook cannot have a Null name!");
        }
        ExcelSheet foundSheet = sheets.stream().filter(excelSheet -> excelSheet.getName().equalsIgnoreCase(name))
                .findFirst().orElse(null);
        if (foundSheet == null) {
            throw new NullPointerException("Workbook doesn't have sheet by the name '" + name + "'!");
        }
        return foundSheet;
    }

    /**
     * Method to create a new Sheet within the Workbook.
     *
     * @param name - Name of the Sheet to create.
     * @throws SheetNotValidException - If the Sheet Creation process fails.
     * @throws IllegalArgumentException - If the name provided isn't valid.
     */
    public ExcelSheet createSheet(String name) throws SheetNotValidException, IllegalArgumentException {
        if (name == null || name.isEmpty()) {
            throw new IllegalArgumentException("Name cannot be invalid!");
        }
        Sheet newSheet = workbook.createSheet(name);
        ExcelSheet sheet = new ExcelSheet(newSheet, true);
        sheets.add(sheet);
        System.out.println("Created Sheet '" + name + "'!");
        return sheet;
    }

    /**
     * Method to return the amount of Sheets loaded from the Workbook.
     *
     * @return - amount of sheets.
     */
    public int getSheetCount() {
        return sheets.size();
    }

    /**
     * Method to return the 'Workbook' hashcode.
     *
     * @return - Bridged HashCode.
     */
    @Override
    public int hashCode() {
        return workbook.hashCode();
    }

    /**
     * Method to determine if any two Objects are equal.
     * Will return false if the Object given is Null or if it isn't an ExcelDocument.
     *
     * @param obj - Any specified Object.
     * @return - false = Not Equal, true = Equal.
     */
    @Override
    public boolean equals(Object obj) {
        if (obj == null) return false;
        if (!(obj instanceof ExcelDocument)) return false;
        ExcelDocument other = (ExcelDocument) obj;
        return workbook.equals(other.getWorkbook());
    }

    /**
     * Method to compare two ExcelDocuments for sorting.
     * Uses SpreadsheetVersions and then the Amount of Pages.
     *
     * @param o - Other ExcelDocument.
     * @return - Comparison Results, '-1', '0' or '1'
     */
    @Override
    public int compareTo(ExcelDocument o) {
        int cmp = workbook.getSpreadsheetVersion().compareTo(o.getWorkbook().getSpreadsheetVersion());
        return cmp == 0 ? Integer.compare(workbook.getNumberOfNames(), o.getWorkbook().getNumberOfNames()) : cmp;
    }

}
