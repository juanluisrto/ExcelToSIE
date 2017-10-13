import java.io.BufferedWriter;
import java.io.IOException;
import java.text.SimpleDateFormat;
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

    public void print(BufferedWriter writer) throws IOException {
        SimpleDateFormat formatVisma = new SimpleDateFormat("yyyyMMdd");
        Entry e = this;
        if (e.ammount > 0) {
            writer.write("#VER A " + e.verfkNummer + " " + formatVisma.format(e.date) + " \"" + e.name + "_" + e.message + "\"\n" +
                    "{\n" +
                    "   #TRANS " + e.debetKonto + " {} -" + e.ammount + "\n" +
                    "   #TRANS " + e.kreditKonto + " {}  " + e.ammount + "\n" +
                    "}\n");
        } else {
            Double opp = e.ammount * -1;
            writer.write("#VER A " + e.verfkNummer + " " + formatVisma.format(e.date) + " \"" + e.name + "_" + e.message + "\"\n" +
                    "{\n" +
                    "   #TRANS " + e.debetKonto + " {} " + opp + "\n" +
                    "   #TRANS " + e.kreditKonto + " {}  " + e.ammount + "\n" +
                    "}\n");
        }

    }
}
