import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Created by juanl on 05/06/2017.
 */
public class Test {
    private static final String EXCEL_FILE_LOCATION = "C:\\Users\\juanl\\Documents\\Secretaria\\verifikationerJavi.xlsx";


    public static void main(String[] args) throws IOException {
        XSSFWorkbook workbook = null;
        try {
            File f = new File(EXCEL_FILE_LOCATION);
            workbook = new XSSFWorkbook(f);

            XSSFSheet sheet = workbook.getSheetAt(0);
            XSSFRow r = sheet.getRow(1);


        SimpleDateFormat formatVisma = new SimpleDateFormat("yyyyMMdd");
        String dateVisma = formatVisma.format(new Date());
        BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(
                new FileOutputStream("verifikationerJavi.SI"), "ibm-437"));
        //PrintWriter writer = new PrintWriter(fileName, "UTF-8");
        writer.write("#FLAGGA 0\n" + //0 significa no importado todavía
                "#FORMAT PC8\n" +
                "#SIETYP 4\n" +
                "#PROGRAM \"Visma Förening\" 2017.0\n" +
                "#GEN " + dateVisma + "\n" +
                "#FNAMN Lärkstaden\n" +
                "#KPTYP EUBAS97\n");
        while(r!=null){
            int i = r.getRowNum() + 21;
            Date date = r.getCell(colNum("B")).getDateCellValue();
            Double debet = r.getCell(colNum("C")).getNumericCellValue();
            Double kredit = r.getCell(colNum("D")).getNumericCellValue();
            String text = r.getCell(colNum("E")).getStringCellValue();
            Double belopp = Double.parseDouble(r.getCell(colNum("F")).getStringCellValue().replace(".", "").replace(",", ".").replace("'", "").replace("\u00A0", "").trim());
            writer.write("#VER A " + i + " " + formatVisma.format(date) + " \"" + text + "\"\n" +
                    "{\n" +
                    "   #TRANS " + debet.intValue() + " {}  " + belopp + "\n" +
                    "   #TRANS " + kredit.intValue()  + " {} -" + belopp + "\n" +
                    "}\n");
            r = sheet.getRow(r.getRowNum() + 1);
        }
        writer.close();
        workbook.close();
            System.out.println("Done");
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }

    }
    public static int colNum(String letter) {
        return CellReference.convertColStringToIndex(letter);
    }
}
