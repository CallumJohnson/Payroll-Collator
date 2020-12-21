package ac.uk.zpq19yru.exceptions;

/*
    
    Created By:     Callum Johnson
    Created In:     Dec/2020
    Project Name:   Payroll Collator
    Package Name:   ac.uk.zpq19yru.exceptions
    Class Purpose:  Exception to be thrown when a workbook is Null.
    
*/

public class WorkbookNotValidException extends Exception {

    /**
     * Constructor to initialise a custom exception.
     *
     * @param message - Message to be sent to 'System.err'.
     */
    public WorkbookNotValidException(String message) {
        super(message);
    }

    /**
     * Constructor to initialise a custom exception.
     */
    public WorkbookNotValidException() {
        super();
    }

}
