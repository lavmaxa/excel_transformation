import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;


import java.io.*;
import java.util.Scanner;

import static org.apache.poi.ss.usermodel.CellType.*;

public class Checking {

    private static Integer getCellWidth(HSSFSheet sheet, HSSFCell cell) {
        int numberOfMergedRegions = sheet.getNumMergedRegions();

        for (int i = 0; i < numberOfMergedRegions; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (firstRow <= cell.getRowIndex() && lastRow >= cell.getRowIndex()) {
                if (cell.getColumnIndex() >= firstColumn && cell.getColumnIndex() <= lastColumn)
                    return lastColumn - firstColumn + 1;
            }
        }
        return 1;
    }

    private static Integer getCellHeight(HSSFSheet sheet, HSSFCell cell) {
        int numberOfMergedRegions = sheet.getNumMergedRegions();

        for (int i = 0; i < numberOfMergedRegions; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (firstRow <= cell.getRowIndex() && lastRow >= cell.getRowIndex()) {
                if (cell.getColumnIndex() >= firstColumn && cell.getColumnIndex() <= lastColumn)
                    return lastRow - firstRow + 1;
            }
        }
        return 1;
    }

    private static String getMergedValue(HSSFSheet sheet, HSSFCell cell) {
        int numberOfMergedRegions = sheet.getNumMergedRegions();

        for (int i = 0; i < numberOfMergedRegions; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (firstRow <= cell.getRowIndex() && lastRow >= cell.getRowIndex()) {
                if (cell.getColumnIndex() >= firstColumn && cell.getColumnIndex() <= lastColumn) {
                    HSSFRow row = sheet.getRow(firstRow);
                    HSSFCell fCell = row.getCell(firstColumn);
                    if (fCell == null) return "";
                    if (fCell.getCellType() == STRING) {
                        return fCell.getStringCellValue();
                    } else if (fCell.getCellType() == BOOLEAN) {
                        return String.valueOf(fCell.getBooleanCellValue());
                    } else if (fCell.getCellType() == FORMULA) {
                        return fCell.getCellFormula();
                    } else if (fCell.getCellType() == NUMERIC) {
                        return String.valueOf(fCell.getNumericCellValue());
                    }
                }
            }
        }
        return "";
    }

    private static String getCellValue(HSSFSheet sheet, HSSFCell fCell) {

        if (fCell == null) return "";

        if (isInMergedRegion(sheet, fCell))
            return getMergedValue(sheet, fCell);
        if (fCell.getCellType() == STRING) {

            return fCell.getStringCellValue();

        } else if (fCell.getCellType() == BOOLEAN) {

            return String.valueOf(fCell.getBooleanCellValue());

        } else if (fCell.getCellType() == FORMULA) {

            return fCell.getCellFormula();

        } else if (fCell.getCellType() == NUMERIC) {

            return String.valueOf(fCell.getNumericCellValue());

        }
        return "";
    }

    private static boolean isInMergedRegion(HSSFSheet sheet, HSSFCell cell) {

        int rb = cell.getRowIndex();
        int re = cell.getRowIndex();
        int cb = cell.getColumnIndex();
        int ce = cell.getColumnIndex();
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (rb >= firstRow && rb <= lastRow || re >= firstRow && re <= lastRow) {
                if (cb >= firstColumn && cb <= lastColumn || ce >= firstColumn && ce <= lastColumn) {
                    return true;
                }
            }
        }
        return false;
    }

    private static String getFileExtension(String mystr) {
        int index = mystr.indexOf('.');
        return index == -1 ? null : mystr.substring(index);
    }

    private static boolean compar(HSSFCell cell1, HSSFCell cell2) {
        HSSFSheet sh1 = cell1.getSheet();
        HSSFSheet sh2 = cell2.getSheet();
        DataFormatter formatter = new DataFormatter();
        FormulaEvaluator evaluator1 = sh1.getWorkbook().getCreationHelper().createFormulaEvaluator();
        FormulaEvaluator evaluator2 = sh2.getWorkbook().getCreationHelper().createFormulaEvaluator();

        if (getCellWidth(sh1, cell1) == getCellWidth(sh2, cell2)) {
            if (getCellHeight(sh1, cell1) == getCellHeight(sh2, cell2)) {
                String d1 = "", d2 = "";
                if (cell1.getCellType() == FORMULA) {
                    d1 = formatter.formatCellValue(cell1, evaluator1);
                }
                if (cell2.getCellType() == FORMULA) {
                    d2 = formatter.formatCellValue(cell2, evaluator2);
                }
                if (getCellValue(sh1, cell1).equals(getCellValue(sh2, cell2)) || d1.equals(d2)) {
                    return true;
                }
            }
        }
        return false;
    }

    public Checking(String d1, String d2, String d3, String d4) throws IOException {

        File[] pathnames;
        File f = new File(d1);
        pathnames = f.listFiles();
        double perc = 0;
        int amount=0;
        File[] ff_res;
        File f1 = new File(d4);
        ff_res = f1.listFiles();
        System.out.println(d1);
        System.out.println(d2);
        System.out.println(d3);
        System.out.println(d4);
        for (File pathname : pathnames) {
            if (getFileExtension(pathname.getName()).equals(".xls")) {
                POIFSFileSystem base = new POIFSFileSystem(new FileInputStream(pathname));
                File expert = new File(d3 + "/" + pathname.getName());
                File result = new File(d2 + "/" + pathname.getName());
                POIFSFileSystem exp = new POIFSFileSystem(new FileInputStream(expert));
                POIFSFileSystem ans = new POIFSFileSystem(new FileInputStream(result));
                HSSFWorkbook wb = new HSSFWorkbook(base);
                HSSFWorkbook wb_exp = new HSSFWorkbook(exp);
                HSSFWorkbook wb_res = new HSSFWorkbook(ans);
                //System.out.println(pathname.getName());
                for (int s = 0; s < wb.getNumberOfSheets(); s++) {
                    HSSFSheet sheet = wb.getSheet(wb.getSheetName(s));
                    HSSFSheet sheet_exp = wb_exp.getSheet(wb.getSheetName(s));
                    HSSFSheet sheet_res = wb_res.getSheet(wb.getSheetName(s));
                    for (File res : ff_res) {
                        int h_beg = -1, h_end = -1, d_beg = -1, d_end = -1;
                        String res_name = pathname.getName() + "____" + sheet.getSheetName();
                        if (res.getName().equals(res_name)) {
                            FileReader reader = new FileReader(d4 + "/" + res_name);
                            Scanner scan = new Scanner(reader);
                            boolean flag = false;
                            while (scan.hasNextLine()) {
                                String cur = scan.nextLine();
                                String str = cur;
                                String numberOnly = str.replaceAll("[^0-9]", "");
                                if (cur.contains("Header")) {
                                    if (flag) {
                                        h_beg = Integer.valueOf(numberOnly);
                                        flag = false;
                                    } else {
                                        if (h_beg == -1) {
                                            h_beg = Integer.valueOf(numberOnly);
                                        } else {
                                            h_end = Integer.valueOf(numberOnly);
                                        }
                                    }
                                }
                                if (cur.contains("Data")) {
                                    if (d_beg == -1) {
                                        d_beg = Integer.valueOf(numberOnly);
                                    } else {
                                        d_end = Integer.valueOf(numberOnly);
                                    }
                                }
                                if (cur.contains("Footnote")) {
                                    flag = true;
                                }
                            }
                            if (d_beg - h_end > 1)
                                h_end = d_beg - 1;
                            if (h_beg != -1 && d_beg != -1) {

                                HSSFRow row = sheet.getRow(h_end - 1);
                                HSSFCell cell = row.getCell(0);
                                if (cell == null)
                                    cell = row.createCell(0);
                                HSSFCellStyle style = cell.getCellStyle();

                                int r = 0;
                                int c = 0;
                                int width = 0;
                                while (style.getBorderBottom() != BorderStyle.THIN && r < 65536) {
                                    row = sheet.getRow(r);
                                    if (row == null) {
                                        row = sheet.createRow(r);
                                    }
                                    cell = row.getCell(c);
                                    if (cell == null) {
                                        cell = row.createCell(c);
                                    }
                                    style = cell.getCellStyle();
                                    r++;
                                }
                                if (r != 65536) {
                                    while (style.getBorderBottom() == BorderStyle.THIN) {
                                        cell = row.getCell(c);
                                        if (cell == null) {
                                            cell = row.createCell(c);
                                        }
                                        style = cell.getCellStyle();
                                        width++;
                                        c++;
                                    }
                                    width--;
                                    //System.out.println(sheet.getSheetName());
                                    int amount_base = 0;
                                    int amount_res = 0;
                                    int mistake = 0;
                                    int amount_exp = 0;
                                    amount_base = sheet.getNumMergedRegions();
                                    amount_exp = sheet_exp.getNumMergedRegions();
                                    amount_res = sheet_res.getNumMergedRegions();
                                    //System.out.println(amount_base + " " + amount_exp + " "  + amount_res);
                                    for (int rr = h_beg - 1; rr < d_end; rr++) {
                                        for (int cc = 0; cc < width; cc++) {
                                            //System.out.println((rr + 1) + " " + (cc + 1));
                                            HSSFRow row1 = sheet.getRow(rr);
                                            if (row1 == null)
                                                row1 = sheet.createRow(rr);
                                            HSSFCell cell1 = row1.getCell(cc);
                                            if (cell1 == null)
                                                cell1 = row1.createCell(cc);
                                            if (getCellHeight(sheet, cell1) == 1 && getCellWidth(sheet, cell1) == 1)
                                                amount_base++;

                                            HSSFRow row2 = sheet_exp.getRow(rr);
                                            if (row2 == null)
                                                row2 = sheet_exp.createRow(rr);
                                            HSSFCell cell2 = row2.getCell(cc);
                                            if (cell2 == null)
                                                cell2 = row2.createCell(cc);
                                            if (getCellHeight(sheet_exp, cell2) == 1 && getCellWidth(sheet_exp, cell2) == 1)
                                                amount_exp++;

                                            HSSFRow row3 = sheet_res.getRow(rr);
                                            if (row3 == null)
                                                row3 = sheet_res.createRow(rr);
                                            HSSFCell cell3 = row3.getCell(cc);
                                            if (cell3 == null)
                                                cell3 = row3.createCell(cc);
                                            if (getCellHeight(sheet_res, cell3) == 1 && getCellWidth(sheet_res, cell3) == 1) {
                                                amount_res++;
                                            }
                                            if (!compar(cell2, cell3))
                                                mistake++;
                                        }
                                    }
                                    amount++;
                                    //System.out.print(amount_base + " " + amount_exp + " "  + amount_res + " " + mistake + " ");
                                    //System.out.println((1 - (double) mistake / (double) amount_exp) * 100);
                                    perc += Math.abs((1 - (double) mistake / (double) amount_exp) * 100);
                                    System.out.println(pathname.getName() + ": " + sheet.getSheetName());
                                }
                            }
                        }
                    }
                }
            }
        }
        System.out.println("AVERAGE PERCENT OF CORRECT TRANSFORMATIONS " + perc / amount);
    }
}