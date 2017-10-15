import java.io.BufferedWriter;
import java.io.IOException;

/**
 * Created by juanl on 13/10/2017.
 */
public abstract class VerificationParent implements Comparable<VerificationParent>{
    int verfkNummer;
    int entryRow;
    boolean exported;
    public abstract void print(BufferedWriter writer) throws IOException;

    public int compareTo(VerificationParent v){
        return (this.verfkNummer > v.verfkNummer) ? 1 : -1;
    }
}
