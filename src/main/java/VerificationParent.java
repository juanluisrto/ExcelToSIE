import java.io.BufferedWriter;
import java.io.IOException;

/**
 * Created by juanl on 13/10/2017.
 */
public abstract class VerificationParent {
    int verfkNummer = 0;


    public abstract void print(BufferedWriter writer) throws IOException;
}
