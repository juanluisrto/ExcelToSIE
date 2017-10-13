import java.util.Date;

/**
 * Created by juanl on 02/06/2017.
 */
public class Entry extends VerificationParent{
    Date date;
    String name;
    String message;
    double ammount;
    String notes;
    int kreditKonto;
    int debetKonto;
    int entryRow;
    //int verfkNummer; inherited
    boolean exported;

    public void print() {

    }
}
