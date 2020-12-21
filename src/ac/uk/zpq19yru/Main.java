package ac.uk.zpq19yru;

/*

    Created By:     Callum Johnson
    Created In:     Dec/2020
    Project Name:   Payroll Collator
    Package Name:   ac.uk.zpq19yru
    Class Purpose:  Main class to run the program.

*/

import ac.uk.zpq19yru.exceptions.WorkbookNotValidException;
import ac.uk.zpq19yru.objects.Configuration;
import ac.uk.zpq19yru.objects.Grade;
import ac.uk.zpq19yru.process.PayrollCollator;

import java.awt.*;
import java.io.Console;
import java.io.File;
import java.io.IOException;

public class Main {

    public static final String SPACER = "=====================================================================";
    public static final String NL_SPACER = "\n" + SPACER + "\n";
    public static final String OUTPUT_FILE_NAME = "UKERLTD - Output.xlsx";
    public static int[] invalid = new int[] {
            14, 16, 18, 19, 20, 21, 22, 23, 24, 25, 26, 28, 29, 30, 31, 33, 34, 35, 36
    };

    public static void main(String[] args) {

        Console console = System.console();
        if(console == null && !GraphicsEnvironment.isHeadless()){
            String filename = Main.class.getProtectionDomain().getCodeSource().getLocation().toString().substring(6)
                    .replaceAll("%20", " ");
            try {
                Runtime.getRuntime().exec(new String[]{"cmd", "/c", "start", "cmd", "/k", "java -jar \"" + filename + "\""});
            } catch (IOException exception) {
                System.out.println("Failed to create Console window.");
            }
            return;
        }

        System.out.println(SPACER);
        System.out.println(
                        "   ___                      _ _     ___      _ _       _             \n" +
                        "  / _ \\__ _ _   _ _ __ ___ | | |   / __\\___ | | | __ _| |_ ___  _ __ \n" +
                        " / /_)/ _` | | | | '__/ _ \\| | |  / /  / _ \\| | |/ _` | __/ _ \\| '__|\n" +
                        "/ ___/ (_| | |_| | | | (_) | | | / /__| (_) | | | (_| | || (_) | |   \n" +
                        "\\/    \\__,_|\\__, |_|  \\___/|_|_| \\____/\\___/|_|_|\\__,_|\\__\\___/|_|   \n" +
                        "            |___/                                                    \n");
        System.out.println(SPACER);
        System.out.println("\nVersion: 1.0");
        System.out.println("Created By Callum Johnson.");
        System.out.println(NL_SPACER);

        try {
            System.out.println("Waiting 5 Seconds before Proceeding.");
            System.out.println(NL_SPACER);
            Thread.sleep(5000);
        } catch (InterruptedException ignored) {
            System.err.println("Could not wait for 5 full seconds, continuing regardless.");
        }

        Configuration properties = new Configuration("rates.properties");

        try {
            properties.loadProperties();
        } catch (IOException exception) {
            System.out.println(
                    "Failed to load properties from the file. Please Contact Callum and DO NOT delete the Output files."
            );
            return;
        }

        PayrollCollator collator = new PayrollCollator();

        System.out.println("Creating Paygrades.");

        for (int i = 1; i <= 6; i++) {
            String gradeId = "grade" + i;
            Grade grade = new Grade(gradeId, properties.getPropertyAsString(gradeId));
            grade.setDailyRate(properties.getPropertyAsDouble(gradeId + "_days"));
            grade.setNightRate(properties.getPropertyAsDouble(gradeId + "_nights"));
            grade.setOTA(properties.getPropertyAsDouble(gradeId + "_ota"));
            grade.setOTB(properties.getPropertyAsDouble(gradeId + "_otb"));
            grade.setBonus(properties.getPropertyAsDouble(gradeId + "_bonus"));
            collator.addPaygrade(grade);
        }

        System.out.println(NL_SPACER);

        System.out.println("Creating the Output Workbook.");
        try {
            collator.createOutputWorkbook();
        } catch (IllegalStateException ex) {
            System.err.println("Failed to create the Output Workbook!");
            return;
        }
        System.out.println("Workbook Successfully Created.");

        System.out.println(NL_SPACER);

        System.out.println("Collecting Documents from the current directory.");
        collator.collectDocuments(new File("."));

        int[] data = collator.amountOfDocumentsFound();
        System.out.println("\nFound " + data[0] + " files, of which " + data[1] + " are relevant.");

        if (data[1] <= 0) {
            System.out.println("\n" + SPACER);
            System.out.println("As there is no relevant files, the process will exit.");
            System.out.println(SPACER);
            return;
        }

        System.out.println(NL_SPACER);

        data = collator.convertDocuments();
        System.out.println("Out of " + data[0] + " documents, " + data[1] + " documents have been converted.");
        if (data[2] > 0) {
            System.out.println(data[2] + " errors were encountered during conversion.");
        }

        if (data[1] <= 0) {
            System.out.println("\n" + SPACER);
            System.out.println("As there is no converted files, the process will exit.");
            System.out.println(SPACER);
            return;
        }

        System.out.println(NL_SPACER);

        System.out.println("Collecting Data from the WorkSheets.");

        collator.collectData();

        System.out.println(NL_SPACER);

        System.out.println("Removing invalid entries from Man Directories. (Columns which don't matter)");
        System.out.println("If a column is removed which is desired, please contact Callum.");

        collator.removeInvalidData();

        System.out.println(NL_SPACER);

        System.out.println("Setting Paygrades using Siemens' Paycodes.");
        collator.setupPaygrades();

        System.out.println(NL_SPACER);

        System.out.println("Outputting Data to the Output File!");
        try {
            collator.outputData();
        } catch (WorkbookNotValidException | IOException | IllegalArgumentException e) {
            e.printStackTrace();
        }

        System.out.println("Saved data to the Output Workbook!");

        System.out.println(NL_SPACER);

        System.out.println("Clearing stored data (inside program)");
        collator.shutdown();
        System.out.println("Cleared. Process Finished.");

        System.out.println(NL_SPACER);

        try {
            System.out.println("Closing this menu in 10 seconds.");
            System.out.println(NL_SPACER);
            Thread.sleep(10000);
        } catch (InterruptedException ignored) {}


    }

}
