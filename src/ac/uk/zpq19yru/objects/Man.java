package ac.uk.zpq19yru.objects;

/*
    
    Created By:     Callum Johnson
    Created In:     Dec/2020
    Project Name:   Payroll Collator
    Package Name:   ac.uk.zpq19yru.objects
    Class Purpose:  Man Object.
    
*/

import ac.uk.zpq19yru.Main;

import java.util.HashMap;
import java.util.Set;

public class Man {

    private final String last;
    private final String first;
    private final HashMap<Integer, Double> data = new HashMap<>();
    private String grade;
    private Grade paygrade;
    private int radiusEntries;

    /**
     * Constructor to initialise a Man Object.
     *
     * @param lastName - Last name of the Man.
     * @param firstName - First name of Man.
     */
    public Man(String lastName, String firstName) {
        this.last = lastName.toUpperCase();
        this.first = firstName.toUpperCase();
    }

    public String getFirst() {
        return first;
    }

    public String getLast() {
        return last;
    }

    public void setGrade(String grade) {
        this.grade = grade;
    }

    public String getGrade() {
        return grade;
    }

    /**
     * Get a result of "Callum Johnson" as a string from Last and First Name variables.
     *
     * @return - 'Firstname Lastname'
     */
    public String getName() {
        return first + " " + last;
    }

    /**
     * Add data to a Man's definition.
     *
     * @param colIndex - Column which the data was found.
     * @param value - Value to be added.
     */
    public void addData(int colIndex, double value) {
        double current = data.getOrDefault(colIndex, 0.0);
        data.put(colIndex, current+value);
        if (colIndex == 27) {
            radiusEntries++;
        }
    }

    /**
     * Return data from Man's definition.
     *
     * @param index - Column of which the data originated.
     * @return - Value or -1 if not found.
     */
    public double getData(int index) {
        if (index == 27) {
            return data.getOrDefault(index, -1.0)/radiusEntries;
        }
        return data.getOrDefault(index, -1.0);
    }

    /**
     * Return all keys in the Data map.
     *
     * @return - Set of Integers.
     */
    public Set<Integer> getKeys() {
        return data.keySet();
    }

    /**
     * Method to remove unrequested data from the listing, these can be changed quite easily for adaptive usage.
     */
    public void removeInvalidEntries() {
        for (int i : Main.invalid) {
            data.remove(i);
        }
    }

    /**
     * Set the Workers' Paygrade (based off of grade)
     *
     * @param grade - Grade to set.
     */
    public void setPayGrade(Grade grade) {
        this.paygrade = grade;
    }

    /**
     * Get the Paygrade or Paygraded Value for the Column Index.
     *
     * @param index - Index of the Paygrade required (Days/Nights/OT1/OT2/Bonus)
     * @param value - Value of the Cell (20 hours worked)
     * @param finalResult - true = 20 * grade, false = grade
     * @return - double value based on 'finalResult'
     */
    public double getPayGrade(int index, double value, boolean finalResult) {
        switch (index) {
            case 12: // Daily
                return finalResult ? paygrade.getDaily()*value : paygrade.getDaily();
            case 13: // Bonus
                return 0;
            case 15: // Travel Hours
                return finalResult ? getRoundOff(value*paygrade.getTravel()) : paygrade.getTravel();
            case 27: // Radius Hours
                return finalResult ? getRoundOff(value*radiusEntries) : value*radiusEntries;
            case 17: // Nights
                return finalResult ? paygrade.getNights()*value : paygrade.getNights();
            case 32: // OTB/OT2
                return finalResult ? paygrade.getOtb()*value : paygrade.getOtb();
            case 37: // OTA/OT1
                return finalResult ? paygrade.getOta()*value : paygrade.getOta();
            default:
                return 0;
        }
    }

    /**
     * Helper method to round-off and calculate a value of a given variable.
     * This method adds NI and other fees onto the value.
     *
     * @param value - Value to be Rounded / Adjusted.
     * @return - final value.
     */
    private double getRoundOff(double value) {
        double eightpointthreethree = value + per(8.33, value);
        double rate = eightpointthreethree + per(13.8, eightpointthreethree);
        return Math.round(rate * 100.0) / 100.0;
    }

    /**
     * Helper method to get the percentage of two figures.
     *
     * @param obtained - Integer, 50 for example.
     * @param total - Integer, 100 for example.
     * @return - Double, 50.0 for example.
     */
    private double per(double obtained, double total) {
        return total * (obtained/100);
    }

    @Override
    public String toString() {
        return "{" + getFirst() + ", " + getLast()
                + " (" + getGrade() + (paygrade != null ? ("/" + paygrade.getName()) : "") + ")" +
                "}";
    }

}
