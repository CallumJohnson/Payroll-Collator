package ac.uk.zpq19yru.process;

/*
    
    Created By:     Callum Johnson
    Created In:     Dec/2020
    Project Name:   Payroll Collator
    Package Name:   ac.uk.zpq19yru.process
    Class Purpose:  Program Instance, used to store all functionality to be called from an object.
    
*/

import ac.uk.zpq19yru.Main;
import ac.uk.zpq19yru.exceptions.CellNotFoundException;
import ac.uk.zpq19yru.exceptions.ManNotFoundException;
import ac.uk.zpq19yru.exceptions.WorkbookNotValidException;
import ac.uk.zpq19yru.objects.ExcelDocument;
import ac.uk.zpq19yru.objects.ExcelSheet;
import ac.uk.zpq19yru.objects.Grade;
import ac.uk.zpq19yru.objects.Man;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.EmptyFileException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class PayrollCollator {

    private int documentsFound = 0, relevantDocumentsFound = 0;
    private final ArrayList<File> relevantFiles = new ArrayList<>();
    private final ArrayList<ExcelDocument> documents = new ArrayList<>();
    private final ArrayList<Man> men = new ArrayList<>();
    private final ArrayList<Grade> grades = new ArrayList<>();
    private File outputFile;
    private XSSFWorkbook outputBook = new XSSFWorkbook();
    private ExcelDocument document;
    private boolean savedHeaders;
    private HashMap<Integer, String> headers = new HashMap<>();

    /**
     * Method to return data collected from the initial scan.
     *
     * @return - { DocumentsFounds, RelevantDocumentsFound }
     */
    public int[] amountOfDocumentsFound() {
        return new int[] { documentsFound, relevantDocumentsFound };
    }

    /**
     * Method to recursively scan a folder to look for XLSX Files.
     *
     * @param file - Folder to scan within.
     */
    public void collectDocuments(File file) {
        System.out.println("Scanning '" + file.getName() + "' for any XLSX files.");
        if (file.listFiles() == null) {
            System.out.println("No Files have been found in the current directory.");
            documentsFound = 0;
            relevantDocumentsFound = 0;
            return;
        }
        for (File subFile : Objects.requireNonNull(file.listFiles())) {
            documentsFound++;
            if (subFile.isDirectory()) {
                collectDocuments(subFile);
                continue;
            }
            String extension = FilenameUtils.getExtension(subFile.getName());
            if (extension.equalsIgnoreCase("xlsx")) {
                relevantDocumentsFound++;
                relevantFiles.add(subFile);
            }
        }
    }

    /**
     * Method to convert all documents into ExcelDocument equivalents.
     *
     * @return - { RelevantDocumentsFound, DocumentsConverted, Errors }.
     */
    public int[] convertDocuments() {
        int documentsConverted = 0, errorsEncountered = 0;
        for (File relevantFile : relevantFiles) {
            InputStream stream;
            try {
                stream = new FileInputStream(relevantFile);
                XSSFWorkbook workbook = new XSSFWorkbook(stream);
                ExcelDocument excelDocument = new ExcelDocument(workbook, relevantFile.getName());
                documents.add(excelDocument);
                documentsConverted++;
            } catch (EmptyFileException ignored) {
            } catch (IOException | WorkbookNotValidException ex) {
                System.err.println(ex.getClass().getSimpleName() + " encountered for file '" + relevantFile.getName() + "'!");
                if (ex.getMessage() != null) {
                    System.err.println("Message Provided: " + ex.getMessage());
                }
                errorsEncountered++;
            }
        }
        return new int[] { relevantDocumentsFound, documentsConverted, errorsEncountered };
    }

    /**
     * Method to collect the data from the ExcelDocuments and their ExcelSheets.
     */
    public void collectData() {
        for (ExcelDocument document : documents) {
            for (int index = 0; index < document.getSheetCount(); index++) {
                ExcelSheet sheet = document.getSheet(index);
                System.out.println("Scanning through sheet '" + sheet.getName()
                        + "' of workbook '" + document.getWorkBookName() + "'!");

                if (!this.savedHeaders) {
                    int row = 0;
                    int maxCols = sheet.getMaxColumns(row);
                    for (int i = 0; i < maxCols; i++) {
                        try {
                            String value = sheet.getCell(row, i).getStringCellValue();
                            headers.put(i, value);
                        } catch (CellNotFoundException e) {
                            break;
                        }
                    }
                    savedHeaders = true;
                    System.out.println("Saved Headers for Output Document.");
                }

                for (int rowNo = 1; rowNo < sheet.getMaxRows(); rowNo++) { // Ignore Header of the Document.
                    int maxColumns = sheet.getMaxColumns(rowNo);
                    if (maxColumns > 2) { // First Name & Last Name required.
                        try {

                            String lastName = sheet.getCell(rowNo, 0).getStringCellValue()
                                    .replaceAll("\n", "");
                            String firstName = sheet.getCell(rowNo, 1).getStringCellValue()
                                    .replaceAll("\n", "");

                            if (firstName.isEmpty() || lastName.isEmpty()) continue;

                            Man man;
                            if (hasManBeenCreated(lastName, firstName)) {
                                try {
                                    man = getManByName(lastName, firstName);
                                } catch (ManNotFoundException ignored) {
                                    System.err.println(
                                            "Despite using the Boolean Method, Man '" + firstName + " " + lastName
                                                    + "' hasn't been found."
                                    );
                                    continue;
                                }
                            } else {
                                man = new Man(lastName, firstName);

                                Cell gradeCell = sheet.getCell(rowNo, 11);
                                man.setGrade(gradeCell.getStringCellValue());

                                men.add(man);
                            }
                            for (int i = 12; i < maxColumns; i++) {
                                Cell cell = sheet.getCell(rowNo, i);
                                man.addData(i, cell.getNumericCellValue());
                            }
                        } catch (CellNotFoundException e) {
                            e.printStackTrace();
                        }
                    }
                }
                System.out.println("Finished Scanning through sheet '" + sheet.getName() + "' of workbook '" + document.getWorkBookName() + "'!");
            }
        }

    }

    /**
     * Determines if a man has already been created and stored into the MEN Array.
     *
     * @param lastName - Last name of the Man.
     * @param firstName - First name of the Man.
     * @return - true = exists, false = doesn't exist.
     */
    private boolean hasManBeenCreated(String lastName, String firstName) {
        return men.stream().anyMatch(
                man -> man.getName().equals(firstName.toUpperCase() + " " + lastName.toUpperCase())
        );
    }

    /**
     * Method to retrieve a Man from their First name and their Last name.
     *
     * @param lastName - Last name of the Man.
     * @param firstName - First name of the Man.
     * @return - Man Object.
     * @throws ManNotFoundException - Thrown if the Man queried doesn't exist.
     */
    private Man getManByName(String lastName, String firstName) throws ManNotFoundException {
        for (Man man : men) {
            if (man.getName().equals(firstName.toUpperCase() + " " + lastName.toUpperCase())) {
                return man;
            }
        }
        throw new ManNotFoundException("Man '" + firstName + " " + lastName + "' not found.");
    }

    /**
     * Method to remove invalid data from ALL Men & header data
     */
    public void removeInvalidData() {
        for (Man man : men) {
            man.removeInvalidEntries();
            System.out.println(man);
        }
        for (int i : Main.invalid) {
            headers.remove(i);
        }
    }

    /**
     * Add a Paygrade to the system.
     *
     * @param grade - Grade to add.
     */
    public void addPaygrade(Grade grade) {
        this.grades.add(grade);
        System.out.println("Added Grade: '" + grade + "'");
    }

    /**
     * Setup PayGrades for men based of their PayCode (Siemens Provided)
     */
    public void setupPaygrades() {
        for (Man man : men) {
            String code = man.getGrade();
            Optional<Grade> grade = grades.stream().filter(g -> g.getLinkedCode().equals(code)).findFirst();
            if (!grade.isPresent()) {
                System.err.println("Code for '" + man.getName() + "' is invalid !!(" + code + ")!!");
                continue;
            }
            man.setPayGrade(grade.get());
            System.out.println(man);
        }
    }

    /**
     * Method to create/clear the output document from local files.
     *
     * @throws IllegalStateException - If the document could not be created.
     */
    public void createOutputWorkbook() throws IllegalStateException {
        this.outputFile = new File(Main.OUTPUT_FILE_NAME);
        if (!this.outputFile.exists()) {
            try {
                if (!this.outputFile.createNewFile()) {
                    System.err.println("Failed to create the output file.");
                    throw new IllegalStateException("Failed to create the output file.");
                }
            } catch (IOException exception) {
                System.err.println("Failed to create the Output File. (IO)");
            }
        } else {
            try {
                FileWriter writer = new FileWriter(this.outputFile);
                writer.write("");
                writer.close();
            } catch (Exception ex) {
                System.err.println("Failed to clear the Output File. (Generic)");
            }
        }
    }

    /**
     * Method to output all of the Data into the Output workbook.
     *
     * @throws WorkbookNotValidException - Thrown if the workbook isn't valid.
     * @throws IllegalArgumentException - Thrown if an argument for the output methods is incorrect.
     */
    public void outputData() throws WorkbookNotValidException, IllegalArgumentException {
        document = new ExcelDocument(outputBook, outputFile.getName());
        ExcelSheet summary = document.createSheet("Summary");
        boolean headersSet = false;
        ExcelSheet hours = document.createSheet("Hours");
        ArrayList<Integer> hoursRequired = new ArrayList<>(Arrays.asList(
                12, 13, 17, 32, 37
        ));
        ExcelSheet expenses = document.createSheet("Expenses");
        ArrayList<Integer> expensesRequired = new ArrayList<>(Arrays.asList(
                15, 27
        ));

        int index = 1;
        int hoursIndex = 0, expensesIndex = 0, columnIndex = 0;
        try {
            for (Man man : men) {
                columnIndex=0;
                hoursIndex=0;
                expensesIndex=0;
                if (!headersSet) {
                    summary.setCell(0, columnIndex++, "Last Name", CellType.STRING, true);
                    hours.setCell(0, hoursIndex++, "Last Name", CellType.STRING, true);
                    expenses.setCell(0, expensesIndex++, "Last Name", CellType.STRING, true);
                    summary.setCell(0, columnIndex++, "First Name", CellType.STRING, true);
                    hours.setCell(0, hoursIndex++, "First Name", CellType.STRING, true);
                    expenses.setCell(0, expensesIndex++, "First Name", CellType.STRING, true);
                    for (Integer key : man.getKeys()) {
                        summary.setCell(0, columnIndex++, headers.get(key), CellType.STRING, true);
                        if (hoursRequired.contains(key)) {
                            hours.setCell(0, hoursIndex++, headers.get(key), CellType.STRING, true);
                        }
                        if (expensesRequired.contains(key)) {
                            expenses.setCell(0, expensesIndex++, headers.get(key), CellType.STRING, true);
                        }
                    }
                    headersSet = true;
                    columnIndex = 0;
                    hoursIndex = 0;
                    expensesIndex = 0;
                }

                summary.setCell(index, columnIndex++, man.getLast(), CellType.STRING, true);
                summary.setCell(index, columnIndex++, man.getFirst(), CellType.STRING, true);

                hours.setCell(index, hoursIndex++, man.getLast(), CellType.STRING, true);
                hours.setCell(index, hoursIndex++, man.getFirst(), CellType.STRING, true);

                expenses.setCell(index, expensesIndex++, man.getLast(), CellType.STRING, true);
                expenses.setCell(index, expensesIndex++, man.getFirst(), CellType.STRING, true);

                boolean hoursUpdate = false, expensesUpdate = false;
                for (Integer key : man.getKeys()) {
                    int temp = index;
                    double value = man.getData(key);
                    summary.setCell(index, columnIndex, value, CellType.NUMERIC, true);
                    if (hoursRequired.contains(key)) {
                        hours.setCell(index, hoursIndex, value, CellType.NUMERIC, true);
                        hoursUpdate = true;
                    }
                    if (expensesRequired.contains(key)) {
                        expenses.setCell(index, expensesIndex, value, CellType.NUMERIC, true);
                        expensesUpdate = true;
                    }
                    index++;
                    summary.setCell(index, columnIndex, man.getPayGrade(key, value, false),
                            CellType.NUMERIC, true);
                    if (hoursRequired.contains(key)) {
                        hours.setCell(index, hoursIndex, man.getPayGrade(key, value, false),
                                CellType.NUMERIC, true);
                    }
                    if (expensesRequired.contains(key)) {
                        expenses.setCell(index, expensesIndex, man.getPayGrade(key, value, false),
                                CellType.NUMERIC, true);
                    }
                    index++;
                    summary.setCell(index, columnIndex, man.getPayGrade(key, value, true),
                            CellType.NUMERIC, true);
                    if (hoursRequired.contains(key)) {
                        hours.setCell(index, hoursIndex, man.getPayGrade(key, value, true),
                                CellType.NUMERIC, true);
                    }
                    if (expensesRequired.contains(key)) {
                        expenses.setCell(index, expensesIndex, man.getPayGrade(key, value, true),
                                CellType.NUMERIC, true);
                    }
                    index = temp;
                    columnIndex++;
                    if (hoursUpdate) {
                        hoursUpdate = false;
                        hoursIndex++;
                    }
                    if (expensesUpdate) {
                        expensesUpdate = false;
                        expensesIndex++;
                    }
                }
                index+=2;
                summary.setCell(
                        index, columnIndex,
                        getFormula(columnIndex, index, hoursRequired.size(), expensesRequired.size(), 1),
                        CellType.FORMULA, true
                );
                hours.setCell(
                        index, hoursIndex,
                        getFormula(hoursIndex, index, hoursRequired.size(), expensesRequired.size(), 2),
                        CellType.FORMULA, true
                );
                expenses.setCell(
                        index, expensesIndex,
                        getFormula(expensesIndex, index, hoursRequired.size(), expensesRequired.size(), 3),
                        CellType.FORMULA, true
                );
                index++;
            }
            summary.setCell(
                    index, columnIndex,
                    getFormula(columnIndex, index, 0,0, 10),
                    CellType.FORMULA, true
            );
            hours.setCell(
                    index, hoursIndex,
                    getFormula(hoursIndex, index, 0,0, 11),
                    CellType.FORMULA, true
            );
            expenses.setCell(
                    index, expensesIndex,
                    getFormula(expensesIndex, index, 0,0, 12),
                    CellType.FORMULA, true
            );
        } catch (Exception ex) {
            System.err.println("Failed to create pages.");
            System.err.println(ex.getClass().getSimpleName() + " has been encountered!");
            if (ex.getMessage() != null) {
                System.err.println(ex.getMessage());
            }
        }

        try {
            FileOutputStream outputStream = new FileOutputStream(outputFile);
            document.getWorkbook().write(outputStream);
            outputStream.close();
            document.getWorkbook().close();
        } catch (IOException ex) {
            System.err.println("Failed to save the workbook.");
        }
    }

    /**
     * Returns a formula to total/sum cells based on location.
     *
     * @param column - Current column.
     * @param row - Current row.
     * @param hoursR - Amount of Hours Cells.
     * @param expensesR - Amount of Expenses Cells.
     * @param page - Page (determines what page you want to return a forumla for).
     * @return - SUM(A1:A2) (example)
     */
    private String getFormula(int column, int row, int hoursR, int expensesR, int page) {
        String formula = "SUM(%BEGIN%:%END%)";
        String begin, last;
        switch (page) {
            case 1:
                begin = convertToCell(column-hoursR-expensesR) + (row+1);
                last = convertToCell(column-1) + (row+1);
                return formula.replaceAll("%BEGIN%", begin).replaceAll("%END%", last);
            case 2:
                begin = convertToCell(column-hoursR) + (row+1);
                last = convertToCell(column-1) + (row+1);
                return formula.replaceAll("%BEGIN%", begin).replaceAll("%END%", last);
            case 3:
                begin = convertToCell(column-expensesR) + (row+1);
                last = convertToCell(column-1) + (row+1);
                return formula.replaceAll("%BEGIN%", begin).replaceAll("%END%", last);
            case 10:
            case 11:
            case 12:
                begin = convertToCell(column) + row;
                last = convertToCell(column) + 2;
                return formula.replaceAll("%BEGIN%", begin).replaceAll("%END%", last);
            default:
                return formula;
        }
    }

    /**
     * Return a Cell Index (A->Z) from an integer.
     *
     * @param index - Integer to convert to A-Z
     * @return - A-Z
     */
    private String convertToCell(int index) {
        String chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        int bigChar = -1;
        if (chars.length() <= index) {
            while (chars.length() <= index) {
                index -= 26;
                bigChar += 1;
            }
            return chars.charAt(bigChar) + "" +  chars.charAt(index);
        } else {
            return "" + chars.charAt(index);
        }
    }

    /**
     * Clear all data centers.
     */
    public void shutdown() {
        this.headers.clear();
        this.documents.clear();
        this.relevantFiles.clear();
        this.men.clear();
        this.grades.clear();
    }

}
