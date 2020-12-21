package ac.uk.zpq19yru.exceptions;

/*
    
    Created By:     Callum Johnson
    Created In:     Dec/2020
    Project Name:   Payroll Collator
    Package Name:   ac.uk.zpq19yru.exceptions
    Class Purpose:  An Exception to be thrown when a Man of work cannot be found.
    
*/

public class ManNotFoundException extends Exception {

    /**
     * Constructor to initialise a custom exception.
     *
     * @param message - Message to be sent to 'System.err'.
     */
    public ManNotFoundException(String message) {
        super(message);
    }

    /**
     * Constructor to initialise a custom exception.
     */
    public ManNotFoundException() {
        super();
    }

}
