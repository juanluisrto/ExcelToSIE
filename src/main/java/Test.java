import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import java.io.File;
import java.io.IOException;

/**
 * Created by juanl on 05/06/2017.
 */
public class Test {
    private static final String EXCEL_FILE_LOCATION = "C:\\Users\\juanl\\Documents\\Secretaria\\verificationer201617.xls";


    public static void main(String[] args) {
        System.out.println("Here I am");
        Workbook workbook = null;
        try {
            System.out.println("Here I am1");
            File f = new File(EXCEL_FILE_LOCATION);
            workbook = Workbook.getWorkbook(f);
            System.out.println("Here I am2");

            Sheet sheet = workbook.getSheet("Verificationer");
            System.out.println(sheet.getCell("G7").getContents());
            System.out.println("H8: " + sheet.getCell("H8").getContents().equals("")); //la celda devuelve un string vacío si no hay na
        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        } finally {

            if (workbook != null) {
                workbook.close();
            }

        }

    }
}
