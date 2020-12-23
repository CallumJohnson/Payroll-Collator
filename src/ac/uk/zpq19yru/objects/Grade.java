package ac.uk.zpq19yru.objects;

/*
    
    Created By:     Callum Johnson
    Created In:     Dec/2020
    Project Name:   Payroll Collator
    Package Name:   ac.uk.zpq19yru.objects
    Class Purpose:  Grade Object.
    
*/

public class Grade {

    private final String name;
    private final String linkedCode;
    private double daily, nights, ota, otb, bonus;
    private double travel;

    /**
     * Constructor to initialise a 'Grade'.
     * A Grade is a payable grade, each grade stands for a different set of Charge Rates.
     *
     * @param name - Name of the Grade, for example 'Grade6'
     * @param linkedCode - Name of the linked Siemens code, for example '073'
     */
    public Grade(String name, String linkedCode) {
        this.name = name;
        this.linkedCode = linkedCode;
    }

    public String getName() {
        return name;
    }

    public String getLinkedCode() {
        return linkedCode;
    }

    public void setDailyRate(double dailyRate) {
        this.daily = dailyRate;
    }

    public void setNightRate(double nightlyRate) {
        this.nights = nightlyRate;
    }

    public void setOTA(double OTA) {
        this.ota = OTA;
    }

    public void setOTB(double OTB) {
        this.otb = OTB;
    }

    public void setTravel(double travel) {
        this.travel = travel;
    }

    @Deprecated
    public void setBonus(double bonus) {
        this.bonus = bonus;
    }

    @Deprecated
    public double getBonus() {
        return bonus;
    }

    public double getTravel() {
        return travel;
    }

    public double getDaily() {
        return daily;
    }

    public double getNights() {
        return nights;
    }

    public double getOta() {
        return ota;
    }

    public double getOtb() {
        return otb;
    }

    @Override
    public String toString() {
        return "Grade{"
                + "name='" + name
                + '\'' + ", linkedCode='" + linkedCode
                + '\'' + ", daily=" + daily
                + ", nights=" + nights
                + ", ota=" + ota
                + ", otb=" + otb
                + ", travel=" + travel
                + ", bonus=" + bonus
                + '}';
    }

}
