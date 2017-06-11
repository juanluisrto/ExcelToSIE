/**
 * Created by juanl on 02/06/2017.
 */

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.*;
import jxl.write.Number;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

public class ExcelToVisma {

    private static final String EXCEL_FILE_LOCATION = "C:\\Users\\juanl\\Documents\\Secretaria\\verificationer201617.xls";
    private static final String EMPTY = "";
    public static ArrayList<Entry> entries = new ArrayList<Entry>();
    public static ArrayList<Rule> rules = new ArrayList<Rule>();
    public static HashMap<Integer, String> benamningar = new HashMap<Integer, String>();
    public static ArrayList<String> nlark = new ArrayList<String>(); //numerarios larkstaden
    public static ArrayList<String> nabr = new ArrayList<String>();  //numerarios åbrink
    public static ArrayList<String> inne = new ArrayList<String>();  //residentes
    public static WritableWorkbook copy;
    public static  Workbook workbook;

    public static void main(String[] args) {

        workbook = null;
        copy = null;
        try {
            File f = new File(EXCEL_FILE_LOCATION);
            workbook = Workbook.getWorkbook(f);
            Sheet sheet = workbook.getSheet("Verificationer");
            importEntries(sheet);
            sheet = workbook.getSheet("Personer");
            importGroups(sheet);
            importRules(sheet);
            sheet = workbook.getSheet("Benamningar");
            importBenamning(sheet);

            System.out.println(rules.toString());
            System.out.println(benamningar.toString());
            File copyfile = new File(EXCEL_FILE_LOCATION.replaceFirst("201617", "201617copy"));
            copy = Workbook.createWorkbook(copyfile,workbook);
            predictEntries();
            copy.write();
            copy.close();
            workbook.close();
            //copy.setOutputFile(f);


        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        } finally {

            if (workbook != null) {
                workbook.close();
            }

        }

    }

    public static void importEntries(Sheet sheet) {
        int row = 7;
        while (!sheet.getCell("H" + row).getContents().equals("END")) {
            if (sheet.getCell("H" + row).getContents().equals(EMPTY)) { //if there is no prediction already, import entry
                Entry e = new Entry();
                e.date = sheet.getCell("A" + row).getContents();
                e.name = sheet.getCell("B" + row).getContents();
                e.message = sheet.getCell("D" + row).getContents();
                e.ammount = Float.parseFloat(sheet.getCell("F" + row).getContents().replace(".", "").replace(",", ".").replace("'", "").replace("�", "").trim());
                e.notes = sheet.getCell("E" + row).getContents();
                e.kontoCell = "I" + row;
                entries.add(e);
            }
            row++;
        }

    }

    public static void importGroups(Sheet sheet) {
        Cell c1 = sheet.findCell("n Larkstaden");
        int row = c1.getRow() + 2 + 1;
        while (sheet.getCell("C" + row).getContents() != EMPTY) {
            nlark.add(sheet.getCell("C" + row).getContents()); //adds all the search queries to nlark set
            row++;
        }
        Cell c2 = sheet.findCell("n Abrink");
        row = c2.getRow() + 2;
        while (sheet.getCell("C" + row).getContents() != EMPTY) {
            nabr.add(sheet.getCell("C" + row).getContents()); //adds all the search queries to nabr set
            row++;
        }
        Cell c3 = sheet.findCell("residentes_Larkstaden");
        row = c3.getRow() + 2;
        while (sheet.getCell("C" + row).getContents() != EMPTY) {
            inne.add(sheet.getCell("C" + row).getContents()); //adds all the search queries to inne set
            row++;
        }
    }

    public static void importRules(Sheet sheet) {
        Cell rulesCell = sheet.findCell("Reglas");
        int row = rulesCell.getRow() + 2;
        int col = rulesCell.getColumn();
        while (sheet.getCell(col, row).getContents() != EMPTY || sheet.getCell(col + 2, row).getContents() != EMPTY || sheet.getCell(col + 3, row).getContents() != EMPTY) {
            //The conditions check if there is a new rule by checking the name, message and ammount fields
            String name = sheet.getCell(col, row).getContents();
            String message = sheet.getCell(col + 2, row).getContents();
            String and = sheet.getCell(col + 1, row).getContents();
            int ammount = -1;
            int margin = -1;
            try { //column-1 row-1
                ammount = Integer.parseInt(sheet.getCell(col + 3, row).getContents());
                margin = Integer.parseInt(sheet.getCell(col + 4, row).getContents());
            } catch (Exception e) {
               // e.printStackTrace();
            }

            int konto = Integer.parseInt(sheet.getCell(col + 6, row).getContents());
            rules.add(new Rule(name, message, and, konto, ammount, margin));
            row++;
        }
    }

    public static void importBenamning(Sheet sheet) {
        int row = 2;
        while (!sheet.getCell(0, row).getContents().equals("END")) {
            if (!sheet.getCell(0, row).getContents().equals(EMPTY)) {
                benamningar.put(Integer.parseInt(sheet.getCell(0, row).getContents()), sheet.getCell(1, row).getContents());
            }
            row++;
        }

    }

    public static void predictEntries() {
        boolean success;
        for (Entry e : entries) {
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
        e.konto = r.konto;
        e.benamning = benamningar.get(r.konto);
        return success;
    }

    private static void writePrediction(Rule r, Entry e) {
        WritableSheet sheet = copy.getSheet("Verificationer");
        WritableCell cell = sheet.getWritableCell(e.kontoCell);
        Label labelCheck = new Label(cell.getColumn() -1, cell.getRow(),"X");
        Label labelKonto = new Label(cell.getColumn(), cell.getRow(), Integer.toString(e.konto));
        Label labelBenamning = new Label(cell.getColumn() + 1, cell.getRow(), e.benamning);
        try {
            sheet.addCell((WritableCell) labelKonto);
            sheet.addCell((WritableCell) labelBenamning);
            sheet.addCell((WritableCell) labelCheck);
        } catch (WriteException e1) {
            e1.printStackTrace();
        }


    }
}

/*


 */