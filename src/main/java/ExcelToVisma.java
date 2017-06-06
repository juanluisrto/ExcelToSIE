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

    private static final String EXCEL_FILE_LOCATION = "C:\\Users\\juanl\\Documents\\Secretaría\\verificationer2016-17.xls";
    private static final String EMPTY = "";
    public static ArrayList<Entry> entries = new ArrayList<Entry>();
    public static ArrayList<Rule> rules = new ArrayList<Rule>();
    public static ArrayList<String> nlark = new ArrayList<String>(); //numerarios larkstaden
    public static ArrayList<String> nabr = new ArrayList<String>();  //numerarios åbrink
    public static ArrayList<String> inne = new ArrayList<String>();  //residentes

    public static void main(String[] args) {

        Workbook  workbook = null;
        try {

            workbook = Workbook.getWorkbook(new File(EXCEL_FILE_LOCATION));
            Sheet sheet = workbook.getSheet("Verificationer");
            importEntries(sheet);
            sheet = workbook.getSheet("Personer");
            importGroups(sheet);
            importRules(sheet);
            predictKonto(sheet);


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

    public static void importEntries(Sheet sheet) {
        int row = 7;
        while (sheet.getCell("H" + row).getContents() != EMPTY) {
            Entry e = new Entry();
            e.date = sheet.getCell("A" + row).getContents();
            e.name = sheet.getCell("B" + row).getContents();
            e.message = sheet.getCell("D" + row).getContents();
            e.ammount = Float.parseFloat(sheet.getCell("F" + row).getContents());
            e.notes = sheet.getCell("E" + row).getContents();
            entries.add(e);
            row++;
        }

    }

    public static void importGroups(Sheet sheet) {
        Cell c1 = sheet.findCell("n Lärkstaden");
        int row = c1.getRow() + 2;
        while (sheet.getCell("C" + row).getContents() != EMPTY) {
            nlark.add(sheet.getCell("C" + row).getContents()); //adds all the search queries to nlark set
            row++;
        }
        Cell c2 = sheet.findCell("n Åbrink");
        row = c2.getRow() + 2;
        while (sheet.getCell("C" + row).getContents() != EMPTY) {
            nabr.add(sheet.getCell("C" + row).getContents()); //adds all the search queries to nabr set
            row++;
        }
        Cell c3 = sheet.findCell("residentes Lärkstaden");
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
            int ammount = Integer.getInteger(sheet.getCell(col + 3, row).getContents());
            int margin = Integer.getInteger(sheet.getCell(col + 4, row).getContents());
            int konto = Integer.getInteger(sheet.getCell(col + 6, row).getContents());
            rules.add(new Rule(name, message, and, konto, ammount, margin));
            row++;
        }
    }

    public static void predictKonto(Sheet sheet) {
        boolean success;
        for (Entry e : entries) {
            for (Rule r : rules) {
                success = entryRuleComparison(r, e);
                if (success) {
                    writePrediction(r, e, sheet);
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
            messageSuccess = e.name.toLowerCase().contains(s.toLowerCase());
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
        if (thereIsAmmount){
            success = success && ammountSuccess;
        }
        return success;
    }
    private static void writePrediction(Rule r, Entry e, Sheet sheet){
        Cell c = sheet.getCell(e.kontoCell);

    }
}

