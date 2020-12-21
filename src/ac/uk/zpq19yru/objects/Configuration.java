package ac.uk.zpq19yru.objects;

/*
    
    Created By:     Callum Johnson
    Created In:     Dec/2020
    Project Name:   Payroll Collator
    Package Name:   ac.uk.zpq19yru.objects
    Class Purpose:  Stands for all Properties.
    
*/

import java.io.*;
import java.nio.file.FileAlreadyExistsException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Properties;

public class Configuration {

    private Properties properties;
    private final String fileName;

    /**
     * Constructor to initialise a .Properties File.
     *
     * @param fileName - FileName to Output from the Jar and Load in as Properties.
     */
    public Configuration(String fileName) {
        this.fileName = fileName;
        System.out.println("Created Properties Instance '" + fileName + "'!");
    }

    /**
     * Method to get a property from a path.
     *
     * @param property - Property Name.
     * @return - Property Value.
     */
    public String getPropertyAsString(String property) {
        return properties.getProperty(property);
    }

    /**
     * Method to get a property from a path.
     *
     * @param property - Property Name.
     * @return - Property Value.
     */
    public double getPropertyAsDouble(String property) {
        return Double.parseDouble(properties.getProperty(property));
    }

    /**
     * Method to set a property Value, ensure to call 'saveProperties()' after this method.
     *
     * @param property - Property Name.
     * @param value - New property value.
     */
    public void setProperty(String property, String value) {
        properties.setProperty(property, value);
    }

    /**
     * Load the properties from file.
     *
     * @throws IOException - An error occurred whilst reading from the input stream.
     */
    public void loadProperties() throws IOException {
        properties = new Properties();
        loadFileFromJar();
        FileInputStream stream = new FileInputStream(new File("./" + fileName));
        properties.load(stream);
        // properties.list(System.out);
        stream.close();
    }

    /**
     * Helper method to load the resource from the Jar.
     *
     * @throws IOException - If the resource cannot be saved.
     */
    private void loadFileFromJar() throws IOException {
        try {
            InputStream inputStream = getClass().getResourceAsStream("/" + fileName);
            Files.copy(inputStream, Paths.get("./" + fileName));
        } catch (FileAlreadyExistsException ignored) {} // Ignore this exception as its okay to occur.
    }

    /**
     * Method to Save Properties in the map to the LocalFile.
     *
     * @throws IOException - Saves all properties to file
     *                       and clears the properties of
     *                       the Configuration instance.
     */
    public void saveProperties() throws IOException {
        OutputStream outputStream = new FileOutputStream("./" + fileName);
        properties.store(outputStream, "Properties of the Program.");
        outputStream.close();
        properties = null;
    }

}
