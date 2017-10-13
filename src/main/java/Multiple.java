import java.io.BufferedWriter;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;

/**
 * Created by juanl on 13/10/2017.
 */
public class Multiple extends VerificationParent {
    Date date;
    String message;
    //int verfkNummer; inherited
    boolean exported;
    HashMap<Integer, Double> konton;


    public void print(BufferedWriter writer) throws IOException {
        SimpleDateFormat formatVisma = new SimpleDateFormat("yyyyMMdd");

        writer.write("#VER A " + this.verfkNummer + " " + formatVisma.format(this.date) + " \"" + this.message + "\"\n " + "{\n" );
        for (Integer k: konton.keySet()){
            writer.write("#TRANS " + k + " {} " + konton.get(k) + "\n");
        }
        writer.write("}\n");

    }
}