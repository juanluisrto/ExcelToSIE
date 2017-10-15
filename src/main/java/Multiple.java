import java.io.BufferedWriter;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import org.apache.commons.lang3.tuple.Pair;

/**
 * Created by juanl on 13/10/2017.
 */
public class Multiple extends VerificationParent {
    Date date;
    String message;
    ArrayList<Pair<Integer,Double>> konton = new ArrayList<Pair<Integer, Double>>();
    //HashMap<Integer, Double> konton = new HashMap<Integer, Double>();
    //inherited
    //int entryRow;
    //int verfkNummer;
    //boolean exported;



    public void print(BufferedWriter writer) throws IOException {
        SimpleDateFormat formatVisma = new SimpleDateFormat("yyyyMMdd");

        writer.write("#VER A " + this.verfkNummer + " " + formatVisma.format(this.date) + " \"" + this.message + "\"\n " + "{\n" );
        for (Pair k: konton){
            writer.write("   #TRANS " + k.getKey() + " {} " + k.getValue() + "\n");
        }
        writer.write("}\n");

    }
}