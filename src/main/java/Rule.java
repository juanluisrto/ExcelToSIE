import java.util.ArrayList;
import java.util.Arrays;

/**
 * Created by juanl on 03/06/2017.
 */
public class Rule {
    ArrayList<String> name = new ArrayList<String>();
    ArrayList<String> message = new ArrayList<String>(3);
    boolean and = false;
    int kreditKonto;
    int debetKonto;
    Double ammount;
    Double margin;

    public Rule(String name, String wildcards, String and, int kreditKonto, int debetKonto, Double ammount, Double margin) {
        if (wildcards.contains(";")) {
            this.message.addAll(Arrays.asList(wildcards.split(";")));
        } else {
            this.message.add(wildcards);
        }
        this.name.add(name); //We add name to arraylist and then we substitute if it was a group
        if (this.name.contains("nlark")) {
            this.name.remove("nlark");
            for (String n : ExcelToSIE.nlark) {
                this.name.add(n);
            }
        }
        if (this.name.contains("nabr")) {
            this.name.remove("nabr");
            for (String n : ExcelToSIE.nabr) {
                this.name.add(n);
            }
        }
        if (this.name.contains("inne")) {
            this.name.remove("inne");
            for (String n : ExcelToSIE.inne) {
                this.name.add(n);
            }
        }
        if (and.equalsIgnoreCase("and")) {
            this.and = true;
        } else {
            this.and = false;
        }
        this.kreditKonto = kreditKonto;
        this.debetKonto = debetKonto;
        if (ammount !=-1) {
            this.ammount = ammount;
            this.margin = margin;
        } else {
            this.ammount = -1.0;
            this.margin = -1.0;
        }



    }
}
