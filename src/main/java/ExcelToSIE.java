/**
 * Created by juanl on 02/06/2017.
 */

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Scanner;

public class ExcelToSIE {

   private static String EXCEL_FILE_LOCATION = "C:\\Users\\juanl\\Documents\\Secretaria\\verifikationer.xlsx";
   //private static String EXCEL_FILE_LOCATION = "verifikationer.xlsx";


    public static ArrayList<Entry> entries = new ArrayList<Entry>();
    public static ArrayList<Rule> rules = new ArrayList<Rule>();
    public static ArrayList<String> nlark = new ArrayList<String>();
    public static ArrayList<String> nabr = new ArrayList<String>();
    public static ArrayList<String> inne = new ArrayList<String>();
    public static ArrayList<Multiple> multiples = new ArrayList<Multiple>();
    public static XSSFWorkbook workbook;
    public static String fnamn;

    public static void main(String[] args) {
        Scanner keyboard = new Scanner(System.in);
        workbook = null;
        try {
            //File in = new File(EXCEL_FILE_LOCATION);
            FileInputStream input = new FileInputStream(EXCEL_FILE_LOCATION);
            workbook = new XSSFWorkbook(input);
            workbook.setMissingCellPolicy(Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            importGroups();
            importRules();
            importEntries();
            System.out.println("Do you want to predict entries with the set of rules? (y/n)");
            String answer = keyboard.nextLine().toLowerCase();
            if (answer.equals("y")) {
                predictEntries();
            }
            System.out.println("Do you want to export the entries to a SIE file? (y/n)");
            answer = keyboard.nextLine().toLowerCase();
            if (answer.equals("y")) {
                exportVerifikationer();
            }
            input.close();
            FileOutputStream out = new FileOutputStream(EXCEL_FILE_LOCATION);
            workbook.write(out);
            out.close();
            workbook.close();
            System.out.println("Done");


        } catch (IOException e) {
            System.out.println(" Excel file not found. Your file should be in the same folder as the jar program");
            e.printStackTrace();
        }

    }

    public static void importMultiples() {
        XSSFSheet sheet = workbook.getSheet("Multipel");
        XSSFRow row = sheet.getRow(3);
        Multiple m = null;
        Double d;
        while (row!=null && row.getCell(colNum("E"), Row.MissingCellPolicy.RETURN_BLANK_AS_NULL)!=null){
            if (row.getCell(colNum("A"), Row.MissingCellPolicy.RETURN_BLANK_AS_NULL)!=null){
                m = new Multiple();
                m.date = row.getCell(colNum("A")).getDateCellValue();
                d = row.getCell(colNum("B")).getNumericCellValue();
                m.verfkNummer = d.intValue();
                m.exported = row.getCell(colNum("C")).getStringCellValue().equals("X");
                m.message = row.getCell(colNum("D")).getStringCellValue();
                multiples.add(m);
            }
            d = row.getCell(colNum("E")).getNumericCellValue();
            Integer konto = d.intValue();
            //retrieves value from debet
            Double ammount = row.getCell(colNum("F"), Row.MissingCellPolicy.RETURN_BLANK_AS_NULL).getNumericCellValue();
            if (ammount == null){ //if the debet cell is null, then the kredit cell has a value
                   ammount = row.getCell(colNum("G"), Row.MissingCellPolicy.RETURN_BLANK_AS_NULL).getNumericCellValue();
                   ammount = -1*ammount;
            }
            m.konton.put(konto,ammount);
            row = sheet.getRow(row.getRowNum()+1);
        }


        }

    public static void importEntries() {
        XSSFSheet sheet = workbook.getSheet("Verifikationer");
        XSSFRow row = sheet.getRow(3);

        while (row!=null && row.getCell(colNum("A"), Row.MissingCellPolicy.RETURN_BLANK_AS_NULL)!=null){
        //while (!row.getCell(colNum("A")).getCellTypeEnum().equals(CellType.BLANK)) {
            Entry e = new Entry();
            e.date = row.getCell(colNum("A")).getDateCellValue();
            e.name = row.getCell(colNum("B")).getStringCellValue();
            e.message = row.getCell(colNum("D")).getStringCellValue();
            e.notes = row.getCell(colNum("E")).getStringCellValue();
            Cell ammount = row.getCell(colNum("F"));
            if (ammount.getCellTypeEnum().equals(CellType.NUMERIC)) {
                e.ammount = ammount.getNumericCellValue();
            } else {
                String editAm = row.getCell(colNum("F")).getStringCellValue().replace(".", "").replace(",", ".").replace("'", "").replace("\u00A0", "").trim(); //�
                e.ammount = Double.parseDouble(editAm);
            }
            e.exported = row.getCell(colNum("H")).getStringCellValue().equals("X");
            Double d = row.getCell(colNum("I")).getNumericCellValue();
            e.verfkNummer = d.intValue();
            e.entryRow = row.getRowNum();
            d = row.getCell(colNum("J")).getNumericCellValue();
            e.debetKonto = d.intValue();
            d = row.getCell(colNum("K")).getNumericCellValue();
            e.kreditKonto = d.intValue();
            entries.add(e);
            row = sheet.getRow(row.getRowNum() + 1);
        }

    }

    public static void importGroups() {
        XSSFSheet sheet = workbook.getSheet("Regler");
        int i = 7;
        Row r = sheet.getRow(i);

        while (r.getCell(colNum("N"), Row.MissingCellPolicy.RETURN_BLANK_AS_NULL) != null || r.getCell(colNum("Q"), Row.MissingCellPolicy.RETURN_BLANK_AS_NULL) != null || r.getCell(colNum("U"), Row.MissingCellPolicy.RETURN_BLANK_AS_NULL) != null) {
            if (r.getCell(colNum("N"), Row.MissingCellPolicy.RETURN_BLANK_AS_NULL) != null) {
                nlark.add(r.getCell(colNum("N")).getStringCellValue());
            }
            if (r.getCell(colNum("Q"), Row.MissingCellPolicy.RETURN_BLANK_AS_NULL) != null) {
                nabr.add(r.getCell(colNum("Q")).getStringCellValue());
            }
            if (r.getCell(colNum("U"), Row.MissingCellPolicy.RETURN_BLANK_AS_NULL) != null) {
                inne.add(r.getCell(colNum("U")).getStringCellValue());
            }
            i++;
            r = sheet.getRow(i);

        }
    }

    public static void importRules() {
        XSSFSheet sheet = workbook.getSheet("Regler");
        XSSFRow r = sheet.getRow(7);
        fnamn = sheet.getRow(2).getCell(colNum("E")).getStringCellValue();
        while (r.getCell(colNum("A"), Row.MissingCellPolicy.RETURN_BLANK_AS_NULL) != null || r.getCell(colNum("C"), Row.MissingCellPolicy.RETURN_BLANK_AS_NULL) != null || r.getCell(colNum("D"), Row.MissingCellPolicy.RETURN_BLANK_AS_NULL) != null) {
            String name = r.getCell(colNum("A")).getStringCellValue();
            String and = r.getCell(colNum("B")).getStringCellValue();
            String message = r.getCell(colNum("C")).getStringCellValue();
            Double ammount = r.getCell(colNum("D")).getNumericCellValue();
            Double margin = r.getCell(colNum("E")).getNumericCellValue();
            Double d = r.getCell(colNum("G")).getNumericCellValue();
            int debet = d.intValue();
            d = r.getCell(colNum("H")).getNumericCellValue();
            int kredit = d.intValue();
            rules.add(new Rule(name, message, and, kredit, debet, ammount, margin));
            r = sheet.getRow(r.getRowNum() + 1);
        }
    }

    public static void predictEntries() {
        boolean success;
        for (Entry e : entries) {
            if (e.debetKonto != 0) {
                continue;
            }
            outerloop:
            for (Rule r : rules) {
                success = entryRuleComparison(r, e);
                if (success) {
                    writePrediction(r, e);
                    break outerloop;
                }
            }
        }
    }

    private static boolean entryRuleComparison(Rule r, Entry e) {
        boolean success = false;
        boolean nameSuccess = false;
        boolean messageSuccess = false;
        boolean thereIsAmmount = r.ammount != 0;
        boolean ammountSuccess = false;
        for (String s : r.name) {
            nameSuccess = e.name.toLowerCase().contains(s.toLowerCase());
            if (s.equals("")) {
                nameSuccess = false;
            }
            if (nameSuccess) {
                break;
            }
        }

        for (String s : r.message) {
            messageSuccess = e.message.toLowerCase().contains(s.toLowerCase());
            if (messageSuccess) {
                break;
            }
        }
        if (thereIsAmmount) {
            ammountSuccess = (e.ammount < (r.ammount + r.margin)) && (e.ammount > (r.ammount - r.margin));
        }
        if (r.and) {
            success = nameSuccess && messageSuccess;
        } else {
            success = nameSuccess || messageSuccess;
        }
        if (thereIsAmmount) {
            success = success && ammountSuccess;
        }
        if (success) {
            e.debetKonto = r.debetKonto;
            e.kreditKonto = r.kreditKonto;
        }
        return success;
    }

    private static void writePrediction(Rule r, Entry e) {
        XSSFSheet sheet = workbook.getSheet("Verifikationer");
        XSSFRow row = sheet.getRow(e.entryRow);
        Cell debetCell = row.createCell(colNum("J"));
        Cell kreditCell = row.createCell(colNum("K"));
        debetCell.setCellValue(r.debetKonto);
        kreditCell.setCellValue(r.kreditKonto);

    }

    public static void exportVerifikationer() {
        XSSFSheet sheet = workbook.getSheet("Verifikationer");
        ArrayList<Entry> toRemove = new ArrayList<Entry>();
        for (Entry e : entries) {
            if (e.debetKonto == 0 || e.exported == true) {
                toRemove.add(e); //we remove the entries without prediction and the ones already exported
            } else {
                XSSFRow row = sheet.getRow(e.entryRow);
                Cell checkCell = row.createCell(colNum("H"));
                checkCell.setCellValue("X");
            }

        }
        entries.removeAll(toRemove);
        //SIE fil format: http://www.sie.se/wp-content/uploads/2014/01/SIE_filformat_ver_4B_080930.pdf
        //Teckenrepertoaren i filen ska vara IBM PC 8-bitars extended ASCII (Codepage 437)
        SimpleDateFormat formatVisma = new SimpleDateFormat("yyyyMMdd");
        String dateVisma = formatVisma.format(new Date());
        String fileName = "verifikationer" + entries.size() + ".SI";
        try {
            BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(
                    new FileOutputStream(fileName), "ibm-437"));
            //PrintWriter writer = new PrintWriter(fileName, "UTF-8");
            writer.write("#FLAGGA 0\n" + //0 significa no importado todavía
                    "#FORMAT PC8\n" +
                    "#SIETYP 4\n" +
                    "#PROGRAM \"Visma Förening\" 2017.0\n" +
                    "#GEN " + dateVisma + "\n" +
                    "#FNAMN " + fnamn + "\n" +
                    "#KPTYP EUBAS97\n");

            for (int i = 0; i < entries.size(); i++) {
                Entry e = entries.get(i);
                //plus för debet och minus för kredit
               e.print(writer);
            }

            writer.close();
        } catch (IOException e) {
        }

    }

    public static int colNum(String letter) {
        return CellReference.convertColStringToIndex(letter);
    }

    public static String numCol(int number) {
        return CellReference.convertNumToColString(number);
    }

}

