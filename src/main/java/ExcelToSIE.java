/**
 * Created by juanl on 02/06/2017.
 */


import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
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
import java.util.HashMap;
import java.util.Iterator;
import java.util.Scanner;

public class ExcelToSIE {

    private static String EXCEL_FILE_LOCATION = "C:\\Users\\juanl\\Documents\\Secretaria\\verifikationer.xlsx";


    public static ArrayList<Entry> entries = new ArrayList<Entry>();
    public static ArrayList<Rule> rules = new ArrayList<Rule>();
    public static HashMap<Integer, String> benamningar = new HashMap<Integer, String>();
    public static ArrayList<String> nlark = new ArrayList<String>(); //numerarios larkstaden
    public static ArrayList<String> nabr = new ArrayList<String>();  //numerarios åbrink
    public static ArrayList<String> inne = new ArrayList<String>();  //residentes
    public static XSSFWorkbook workbook;

    public static void main(String[] args) {
        Scanner keyboard = new Scanner(System.in);
        /*System.out.println("Enter the name of your .xlsx file (default= verificationer.xlsx): ");
        EXCEL_FILE_LOCATION = keyboard.next();*/
        workbook = null;
        try {
            File in = new File(EXCEL_FILE_LOCATION);
            FileOutputStream out = new FileOutputStream(EXCEL_FILE_LOCATION + "out");
            workbook = new XSSFWorkbook(in);
            workbook.setMissingCellPolicy(Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            importGroups();
            importRules();
            importBenamning();
            importEntries();
            System.out.println(rules.toString());
            //System.out.println(benamningar.toString());
            //predictEntries();
            System.out.println("Do you want to export the entries to a SIE file? (y/n)");
            String answer = keyboard.nextLine().toLowerCase();
            if (answer.equals("y")) {
                exportVerificationer();
            }
            workbook.write(out);
            out.close();
            workbook.close();
            System.out.println("Done");


        } catch (IOException e) {
            System.out.println("File not found. Your file should be in the same folder as the jar program");
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }

    }

    public static void importEntries() {
        XSSFSheet sheet = workbook.getSheet("Verificationer");
        XSSFRow row = sheet.getRow(6);

        while (row.getCell(colNum("H")) == null || !row.getCell(colNum("H")).getStringCellValue().equals("END")) {
            if (row.getCell(colNum("H")) == null) { //if verification has not been exported yet (cell=null), import entry to be processed
                Entry e = new Entry();
                e.date = row.getCell(colNum("A")).getDateCellValue();
                e.name = row.getCell(colNum("B")).getStringCellValue();
                e.message = row.getCell(colNum("D")).getStringCellValue();
                String editAm = row.getCell(colNum("F")).getStringCellValue().replace(".", "").replace(",", ".").replace("'", "").replace("\u00A0", "").trim(); //�
                e.ammount = Double.parseDouble(editAm);
                e.notes = row.getCell(colNum("E")).getStringCellValue();
                e.entryRow = row.getRowNum();
                if (row.getCell(colNum("I")) != null) { // If there is already a prediction, import it.
                    Double d = row.getCell(colNum("I")).getNumericCellValue();
                    e.debetKonto = d.intValue();
                    e.benamning = benamningar.get(e.debetKonto);
                }
                entries.add(e);
            }
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

        while (r.getCell(colNum("A"), Row.MissingCellPolicy.RETURN_BLANK_AS_NULL) != null || r.getCell(colNum("C"), Row.MissingCellPolicy.RETURN_BLANK_AS_NULL) != null || r.getCell(colNum("D"), Row.MissingCellPolicy.RETURN_BLANK_AS_NULL) != null) {
            String name = r.getCell(colNum("A")).getStringCellValue();
            String and = r.getCell(colNum("B")).getStringCellValue();
            String message = r.getCell(colNum("C")).getStringCellValue();
            Double ammount = r.getCell(colNum("D")).getNumericCellValue();
            Double margin = r.getCell(colNum("E")).getNumericCellValue();
            Double d = r.getCell(colNum("G")).getNumericCellValue();
            int kredit = d.intValue();
            d = r.getCell(colNum("H")).getNumericCellValue();
            int debet = d.intValue();
            rules.add(new Rule(name, message, and, kredit, debet, ammount, margin));
            r = sheet.getRow(r.getRowNum() + 1);


            /*while (!row.getCell(colNum("F"), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).getStringCellValue().equalsIgnoreCase("END")) {
                //while (row.getCell(colNum("F"))!=null || row.getCell(colNum("H"))!=null || row.getCell(colNum("I"))!=null){
                //The conditions check if there is a new rule by checking the name, message and ammount fields //.getRawValue() != ""


                Double ammount;
                Double margin;

                try {
                    ammount = row.getCell(;
                    margin = row.getCell(colNum("J")).getNumericCellValue();
                } catch (Exception e) {
                    ammount = -1.;
                    margin = -1.;
                }


            }*/
        }
    }

    public static void importBenamning() {
        XSSFSheet sheet = workbook.getSheet("Konton");
        XSSFRow row = sheet.getRow(3);
        while (row.getCell(colNum("A")).getCellTypeEnum() != CellType.STRING) {
            if (row.getCell(colNum("A")) != null) {
                Double d = row.getCell(colNum("A")).getNumericCellValue();
                benamningar.put(d.intValue(), row.getCell(colNum("B")).getStringCellValue());
            }
            row = sheet.getRow(row.getRowNum() + 1);
        }
        int i = 0;
    }

    public static void predictEntries() {
        boolean success;
        for (Entry e : entries) {
            if (e.benamning != null) {
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
        boolean thereIsAmmount = r.ammount != -1;
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
            e.benamning = benamningar.get(r.debetKonto);
        }
        return success;
    }

    private static void writePrediction(Rule r, Entry e) {
        XSSFSheet sheet = workbook.getSheet("Verificationer");
        XSSFRow row = sheet.getRow(e.entryRow);
        Cell kontoCell = row.createCell(colNum("I"));
        Cell benamningCell = row.createCell(colNum("J"));
        kontoCell.setCellValue(r.debetKonto);
        benamningCell.setCellValue(e.benamning);

    }

    public static void exportVerificationer() {
        XSSFSheet sheet = workbook.getSheet("Verificationer");
        ArrayList<Entry> toRemove = new ArrayList<Entry>();
        for (Entry e : entries) {
            if (e.benamning == null) {
                toRemove.add(e); //we remove the entries without prediction
            } else {
                XSSFRow row = sheet.getRow(e.entryRow);
                Cell checkCell = row.createCell(colNum("H"));
                checkCell.setCellValue("X");
            }

        }
        entries.removeAll(toRemove);
        //SIE fil format: http://www.sie.se/wp-content/uploads/2014/01/SIE_filformat_ver_4B_080930.pdf
        //Teckenrepertoaren i filen ska vara IBM PC 8-bitars extended ASCII (Codepage 437)
        //String dateSeconds = new SimpleDateFormat("dd-MM HH:mm:ss").format(new Date());
        SimpleDateFormat formatVisma = new SimpleDateFormat("yyyyMMdd");
        String dateVisma = formatVisma.format(new Date());
        String fileName = "C:\\Users\\juanl\\Documents\\Secretaria\\verificationer" + entries.size() + ".SI";
        try {
            BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(
                    new FileOutputStream(fileName), "ibm-437"));
            //PrintWriter writer = new PrintWriter(fileName, "UTF-8");
            writer.write("#FLAGGA 0\n" + //0 significa no importado todavía
                    "#FORMAT PC8\n" +
                    "#SIETYP 4\n" +
                    "#PROGRAM \"Visma Förening\" 2017.0\n" +
                    "#GEN " + dateVisma + "\n" +
                    "#FNAMN Lärkstaden\n" +
                    "#KPTYP EUBAS97\n");
            //Se añaden las cuentas
            writer.write("#KONTO 1010 \"Kassa\"\n" +
                    "#KTYP 1010 T\n" +
                    "#KONTO 1020 \"Plusgirot\"\n" +
                    "#KTYP 1020 T\n" +
                    "#KONTO 1300 \"Interims Fordring\"\n" +
                    "#KTYP 1300 T\n" +
                    "#KONTO 1310 \"Extern Fordring\"\n" +
                    "#KTYP 1310 T\n" +
                    "#KONTO 2010 \"Deposition\"\n" +
                    "#KTYP 2010 S\n" +
                    "#KONTO 2300 \"Interims Skuld\"\n" +
                    "#KTYP 2300 S\n" +
                    "#KONTO 2910 \"Eget Kapital\"\n" +
                    "#KTYP 2910 S\n" +
                    "#KONTO 3008 \"Logi ctr\"\n" +
                    "#KTYP 3008 T\n" +
                    "#DK 3008 \"K\"\n" +
                    "#KONTO 3009 \"Logi åbrink\"\n" +
                    "#KTYP 3009 I\n" +
                    "#DK 3009 \"K\"\n" +
                    "#KONTO 3010 \"Logi inneb\"\n" +
                    "#KTYP 3010 I\n" +
                    "#DK 3010 \"K\"\n" +
                    "#KONTO 3011 \"Boende\"\n" +
                    "#KTYP 3011 I\n" +
                    "#DK 3011 \"K\"\n" +
                    "#KONTO 3012 \"Bidrag åbrink\"\n" +
                    "#KTYP 3012 I\n" +
                    "#DK 3012 \"K\"\n" +
                    "#KONTO 3013 \"Bidrag LSC\"\n" +
                    "#KTYP 3013 I\n" +
                    "#DK 3013 \"K\"\n" +
                    "#KONTO 3020 \"Gäster\"\n" +
                    "#KTYP 3020 I\n" +
                    "#DK 3020 \"K\"\n" +
                    "#KONTO 3040 \"Telefon\"\n" +
                    "#KTYP 3040 I\n" +
                    "#DK 3040 \"K\"\n" +
                    "#KONTO 3110 \"Aktiviteter\"\n" +
                    "#KTYP 3110 I\n" +
                    "#DK 3110 \"K\"\n" +
                    "#KONTO 3120 \"Sommarkurs LSC\"\n" +
                    "#KTYP 3120 I\n" +
                    "#DK 3120 \"K\"\n" +
                    "#KONTO 3121 \"Surahammar\"\n" +
                    "#KTYP 3121 T\n" +
                    "#DK 3121 \"K\"\n" +
                    "#KONTO 3122 \"Sura sr\"\n" +
                    "#KTYP 3122 I\n" +
                    "#DK 3122 \"K\"\n" +
                    "#KONTO 3130 \"Gåvor och Bidrag\"\n" +
                    "#KTYP 3130 I\n" +
                    "#DK 3130 \"K\"\n" +
                    "#KONTO 3140 \"Bilens användarintäkter\"\n" +
                    "#KTYP 3140 I\n" +
                    "#DK 3140 \"K\"\n" +
                    "#KONTO 5010 \"Hyra boende\"\n" +
                    "#KTYP 5010 K\n" +
                    "#KONTO 5015 \"För amorteringar\"\n" +
                    "#KTYP 5015 K\n" +
                    "#KONTO 5020 \"Inredning\"\n" +
                    "#KTYP 5020 K\n" +
                    "#KONTO 5030 \"Telefon abonemang\"\n" +
                    "#KTYP 5030 K\n" +
                    "#KONTO 5031 \"Telefon samtal\"\n" +
                    "#KTYP 5031 K\n" +
                    "#KONTO 5050 \"Ekonomiförvalting\"\n" +
                    "#KTYP 5050 K\n" +
                    "#KONTO 5060 \"Livsmedel\"\n" +
                    "#KTYP 5060 K\n" +
                    "#KONTO 5070 \"Glödlampor\"\n" +
                    "#KTYP 5070 K\n" +
                    "#KONTO 5071 \"Reparationer\"\n" +
                    "#KTYP 5071 K\n" +
                    "#KONTO 5072 \"Tidningar\"\n" +
                    "#KTYP 5072 K\n" +
                    "#KONTO 5073 \"Diverse\"\n" +
                    "#KTYP 5073 K\n" +
                    "#KONTO 6010 \"Aktiviteter\"\n" +
                    "#KTYP 6010 K\n" +
                    "#KONTO 6011 \"Surahammar\"\n" +
                    "#KTYP 6011 K\n" +
                    "#KONTO 6012 \"Sura sr\"\n" +
                    "#KTYP 6012 K\n" +
                    "#KONTO 6015 \"Extra lön (sommarkurs)\"\n" +
                    "#KTYP 6015 K\n" +
                    "#KONTO 6020 \"PR-kostnader\"\n" +
                    "#KTYP 6020 K\n" +
                    "#KONTO 6030 \"Hyra lokaler\"\n" +
                    "#KTYP 6030 K\n" +
                    "#KONTO 6035 \"Telefon\"\n" +
                    "#KTYP 6035 K\n" +
                    "#KONTO 6040 \"Porto\"\n" +
                    "#KTYP 6040 K\n" +
                    "#KONTO 6041 \"Video- och fotokostnader\"\n" +
                    "#KTYP 6041 K\n" +
                    "#KONTO 6042 \"Diverse\"\n" +
                    "#KTYP 6042 K\n" +
                    "#KONTO 6090 \"\u008Frligabilkostnader\"\n" +
                    "#KTYP 6090 K\n" +
                    "#KONTO 6091 \"Bensin\"\n" +
                    "#KTYP 6091 K\n" +
                    "#KONTO 6092 \"Driftkostnader bil\"\n" +
                    "#KTYP 6092 K\n" +
                    "#KONTO 8110 \"PG-avgift\"\n" +
                    "#KTYP 8110 K\n" +
                    "#DIM 1 Resultatenhet\n" +
                    "#DIM 6 Projekt\n");

            for (int i = 0; i < entries.size(); i++) {
                Entry e = entries.get(i);
                //plus för debet och minus för kredit
                if (e.ammount > 0) {
                    writer.write("#VER A " + i + " " + formatVisma.format(e.date) + " \"" + e.name + "_" + e.message + "\"\n" +
                            "{\n" +
                            "   #TRANS " + e.debetKonto + " {} -" + e.ammount + "\n" +
                            "   #TRANS 1020 {} " + e.ammount + "\n" +
                            "}\n");
                } else {
                    Double opp = e.ammount * -1;
                    writer.write("#VER A " + i + " " + formatVisma.format(e.date) + " \"" + e.name + "_" + e.message + "\"\n" +
                            "{\n" +
                            "   #TRANS " + e.debetKonto + " {} " + opp + "\n" +
                            "   #TRANS 1020 {} " + e.ammount + "\n" +
                            "}\n");
                }
            }

            writer.close();
        } catch (IOException e) {
            // do something
        }

    }

    public static int colNum(String letter) {
        return CellReference.convertColStringToIndex(letter);
    }

    public static String numCol(int number) {
        return CellReference.convertNumToColString(number);
    }

}

/*


 */