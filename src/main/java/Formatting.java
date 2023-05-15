import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.commons.lang3.math.NumberUtils;

import java.io.*;
import java.util.ArrayList;
import java.util.Scanner;

import static org.apache.poi.ss.usermodel.CellType.*;

public class Formatting {

    private static void vert_expansion(HSSFSheet sheet, HSSFCell cell, int he) {
        String d = "";
        int cc = cell.getColumnIndex(), ce = cc, cb = cc, rb, re;
        HSSFCellStyle st;
        st = cell.getCellStyle();
        int rr = cell.getRowIndex();
        rb = rr;
        re = rr;
        while (st.getBorderBottom() != BorderStyle.THIN && rr <= he) {
            if (d.equals(""))
                d = getCellValue(sheet, cell);
            HSSFRow row_low = sheet.getRow(rr + 1);
            HSSFCell lowcell = row_low.getCell(cc);
            if (lowcell == null) {
                lowcell = row_low.createCell(cc);
            }
            if (getCellWidth(sheet, lowcell) != getCellWidth(sheet, cell)) {
                break;
            }
            if (!getCellValue(sheet, cell).equals("") && !getCellValue(sheet, lowcell).equals("") && !getCellValue(sheet, cell).equals(getCellValue(sheet, lowcell))) {
                break;
            }
            rr++;
            re = rr;
            HSSFRow row = sheet.getRow(rr);
            cell = row.getCell(cc);
            if (cell == null) {
                cell = row.createCell(cb);
            }
            st = cell.getCellStyle();
        }
        HSSFRow row = sheet.getRow(rb);
        cell = row.getCell(cb);
        if (cell == null) {
            cell = row.createCell(cb);
        }
        ce = cb + getHorSize(sheet, cell) - 1;
        makeNewMerge(sheet, rb, re, cb, ce, d);
    }

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

    private static Integer getVertSize(HSSFSheet sheet, HSSFCell cell) {
        int numberOfMergedRegions = sheet.getNumMergedRegions();

        for (int i = 0; i < numberOfMergedRegions; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (firstRow <= cell.getRowIndex() && lastRow >= cell.getRowIndex()) {
                if (cell.getColumnIndex() >= firstColumn && cell.getColumnIndex() <= lastColumn)
                    return lastRow - cell.getRowIndex() + 1;
            }
        }
        return 1;
    }

    private static Integer getHorSize(HSSFSheet sheet, HSSFCell cell) {
        int numberOfMergedRegions = sheet.getNumMergedRegions();

        for (int i = 0; i < numberOfMergedRegions; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (firstRow <= cell.getRowIndex() && lastRow >= cell.getRowIndex()) {
                if (cell.getColumnIndex() >= firstColumn && cell.getColumnIndex() <= lastColumn)
                    return lastColumn - cell.getColumnIndex() + 1;
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

    private static Integer getMergedWidth(HSSFSheet sheet, HSSFCell cell) {
        int numberOfMergedRegions = sheet.getNumMergedRegions();
        for (int i = 0; i < numberOfMergedRegions; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (firstRow <= cell.getRowIndex() && lastRow >= cell.getRowIndex()) {
                if (cell.getColumnIndex() >= firstColumn && cell.getColumnIndex() <= lastColumn) {
                    return lastColumn - firstColumn + 1;
                }
            }
        }
        return 1;
    }

    private static Integer getMergedHeight(HSSFSheet sheet, HSSFCell cell) {
        int numberOfMergedRegions = sheet.getNumMergedRegions();
        for (int i = 0; i < numberOfMergedRegions; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (firstRow <= cell.getRowIndex() && lastRow >= cell.getRowIndex()) {
                if (cell.getColumnIndex() >= firstColumn && cell.getColumnIndex() <= lastColumn) {
                    return lastRow - firstRow + 1;

                }
            }
        }
        return 1;
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

    private static void putCellValue(HSSFCell cell, String d) {
        if (NumberUtils.isNumber(d)) {
            DataFormat format = cell.getSheet().getWorkbook().createDataFormat();
            HSSFCellStyle st = cell.getCellStyle();
            String j = "";
            for (int i = 0; i < d.length(); i++) {
                if (!(d.charAt(i) >= '0' && d.charAt(i) <= '9')) {
                    break;
                }
                j += d.charAt(i);
            }
            if (d.charAt(0) == '-') {
                for (int i = 1; i < d.length(); i++) {
                    if (!(d.charAt(i) >= '0' && d.charAt(i) <= '9')) {
                        break;
                    }
                    j += d.charAt(i);
                }
            }

            double h = NumberUtils.createDouble(j) * (-1);
            if (Math.abs(Math.floor(h) - Math.ceil(h)) < 0.00000001) {
                st.setDataFormat(HSSFDataFormat.getBuiltinFormat("General"));
                cell.setCellValue(NumberUtils.createInteger(j));

            } else {
                st.setDataFormat(format.getFormat("#############.######"));
                cell.setCellValue(NumberUtils.createDouble(d));
            }
        } else
            cell.setCellValue(d);
    }

    private static void makeNewMerge(HSSFSheet sheet, int rb, int re, int cb, int ce, String d) {
        HSSFRow row4 = sheet.getRow(rb);
        HSSFCell cell4 = row4.getCell(cb);
        if (cell4 == null) {
            cell4 = row4.createCell(cb);
        }
        ArrayList<Integer> ind = new ArrayList<>();
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = sheetMergeCount; i >= 0; i--) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            if (range != null) {
                int firstColumn = range.getFirstColumn();
                int lastColumn = range.getLastColumn();
                int firstRow = range.getFirstRow();
                int lastRow = range.getLastRow();
                if (rb == firstRow && re == lastRow && cb == firstColumn && ce == lastColumn) {
                    return;
                }
                if (ce - cb + 1 <= getMergedWidth(sheet, cell4) && re - rb + 1 <= getMergedHeight(sheet, cell4)) {
                    return;
                }
                if (firstRow >= rb && firstRow <= re && lastRow >= rb && lastRow <= re) {
                    if (firstColumn >= cb && firstColumn <= ce && lastColumn >= cb && lastColumn <= ce) {
                        if (d.equals(""))
                            d = getCellValue(sheet, sheet.getRow(firstRow).getCell(firstColumn));
                        ind.add(i);
                    }
                }
            }
        }
        if (!(rb == re && cb == ce)) {
            for (int i = 0; i < ind.size(); i++) {
                sheet.removeMergedRegion(ind.get(i));
            }
            HSSFRow row3 = sheet.getRow(rb);
            HSSFCell cell3 = row3.getCell(cb);
            if (cell3 == null) {
                cell3 = row3.createCell(cb);
            }
            if (d.equals("")) {
                for (int r = rb; r <= re; r++) {
                    for (int c = cb; c < ce; c++) {
                        if (d.equals(""))
                            d = getCellValue(sheet, sheet.getRow(r).getCell(c));
                    }
                }
            }
            putCellValue(cell3, d);
            CellRangeAddress region = new CellRangeAddress(rb, re, cb, ce);
            try {
                sheet.addMergedRegion(region);
            } catch (Exception e) {
            }
            CellUtil.setVerticalAlignment(cell3, VerticalAlignment.CENTER);
            CellUtil.setAlignment(cell3, HorizontalAlignment.CENTER);
            HSSFCellStyle st;
            st = cell3.getCellStyle();
            st.setWrapText(true);
        } else {
            return;
        }
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

    private static boolean isInOneMerged(HSSFSheet sheet, HSSFCell cell1, HSSFCell cell2) {
        int rb1 = cell1.getRowIndex();
        int re1 = cell1.getRowIndex();
        int cb1 = cell1.getColumnIndex();
        int ce1 = cell1.getColumnIndex();
        int rb2 = cell2.getRowIndex();
        int re2 = cell2.getRowIndex();
        int cb2 = cell2.getColumnIndex();
        int ce2 = cell2.getColumnIndex();
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (rb1 >= firstRow && rb1 <= lastRow || re1 >= firstRow && re1 <= lastRow) {
                if (cb1 >= firstColumn && cb1 <= lastColumn || ce1 >= firstColumn && ce1 <= lastColumn) {
                    if (rb2 >= firstRow && rb2 <= lastRow || re2 >= firstRow && re2 <= lastRow) {
                        if (cb2 >= firstColumn && cb2 <= lastColumn || ce2 >= firstColumn && ce2 <= lastColumn) {
                            return true;
                        }
                    }
                }
            }
        }
        return false;
    }

    private static String getFileExtension(String mystr) {
        int index = mystr.indexOf('.');
        return index == -1 ? null : mystr.substring(index);
    }

    private static void copyFileUsingStream(File source, File dest) throws IOException {
        InputStream is = null;
        OutputStream os = null;
        try {
            is = new FileInputStream(source);
            os = new FileOutputStream(dest);
            byte[] buffer = new byte[1024];
            int length;
            while ((length = is.read(buffer)) > 0) {
                os.write(buffer, 0, length);
            }
        } finally {
            is.close();
            os.close();
        }
    }

    public Formatting(String d1, String d2, String d3) throws IOException {

        File[] pathnames;
        File f = new File(d1);
        pathnames = f.listFiles();

        File[] ff_res;
        File f1 = new File(d2);
        ff_res = f1.listFiles();

        for (File pathname : pathnames) {
            if (getFileExtension(pathname.getName()).equals(".xls")) {

                File ff = new File(d3 + "/" + pathname.getName());
                copyFileUsingStream(pathname, ff);
                POIFSFileSystem findfile = new POIFSFileSystem(new FileInputStream(ff));
                POIFSFileSystem base = new POIFSFileSystem(new FileInputStream(pathname));
                FileOutputStream fileOut = new FileOutputStream(ff);
                HSSFWorkbook wb = new HSSFWorkbook(findfile);
                HSSFWorkbook wb_base = new HSSFWorkbook(base);
                System.out.println(pathname.getName());
                for (int s = 0; s < wb.getNumberOfSheets(); s++) {
                    HSSFSheet sheet = wb.getSheet(wb.getSheetName(s));
                    for (File res : ff_res) {
                        int h_beg = -1, h_end = -1, d_beg = -1, d_end = -1;
                        String res_name = pathname.getName() + "____" + sheet.getSheetName();
                        if (res.getName().equals(res_name)) {
                            FileReader reader = new FileReader(d2 + "/" + res_name);
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
                                //header
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
                                    System.out.println(sheet.getSheetName());
                                    //HEADER Vertical expansion
                                    for (int cc = 0; cc < width; cc++) {
                                        String d = "";
                                        int cb = cc, ce = cc, rb, re;
                                        for (int rr = h_beg - 1; rr < h_end; rr++) {
                                            row = sheet.getRow(rr);
                                            cell = row.getCell(cc);
                                            if (cell == null) cell = row.createCell(cb);
                                            HSSFCellStyle st;
                                            st = cell.getCellStyle();
                                            rb = rr;
                                            re = rr;
                                            while (st.getBorderBottom() != BorderStyle.THIN && rr < h_end) {
                                                if (d.equals(""))
                                                    d = getCellValue(sheet, cell);
                                                HSSFRow row_low = sheet.getRow(rr + 1);
                                                HSSFCell lowcell = row_low.getCell(cc);
                                                if (lowcell == null) lowcell = row_low.createCell(cc);
                                                if (getCellWidth(sheet, lowcell) != getCellWidth(sheet, cell)) {
                                                    break;
                                                }
                                                if (!getCellValue(sheet, cell).equals("") && !getCellValue(sheet, lowcell).equals("") && !getCellValue(sheet, cell).equals(getCellValue(sheet, lowcell))) {
                                                    break;
                                                }
                                                rr++;
                                                re = rr;
                                                row = sheet.getRow(rr);
                                                cell = row.getCell(cc);
                                                if (cell == null) cell = row.createCell(cc);
                                                st = cell.getCellStyle();
                                            }
                                            row = sheet.getRow(rb);
                                            cell = row.getCell(cb);
                                            if (cell == null) cell = row.createCell(cb);
                                            ce = cb + getHorSize(sheet, cell) - 1;
                                            makeNewMerge(sheet, rb, re, cb, ce, d);
                                            d = "";
                                        }
                                    }
                                    //HEADER Horizontal expansion
                                    for (int rr = h_beg - 1; rr < h_end; rr++) {
                                        String d = "";
                                        int ind = -1;
                                        int cb, ce, rb = rr, re = rr;
                                        for (int cc = 0; cc < width; cc++) {
                                            row = sheet.getRow(rr);
                                            cell = row.getCell(cc);
                                            if (cell == null) cell = row.createCell(cc);
                                            HSSFCellStyle st;
                                            st = cell.getCellStyle();
                                            cb = cc;
                                            ce = cc;
                                            while (st.getBorderRight() != BorderStyle.THIN && cc < width - 1) {
                                                if (d.equals("")) {
                                                    d = getCellValue(sheet, cell);
                                                    if (!getCellValue(sheet, cell).equals(""))
                                                        ind = cc;
                                                } else if (!getCellValue(sheet, cell).equals("") && !getCellValue(sheet, cell).equals(d)) {
                                                    cc--;
                                                    break;
                                                } else if (getCellValue(sheet, cell).equals("")) {
                                                    cc--;
                                                    break;
                                                }
                                                HSSFCell rc = row.getCell(cc + 1);
                                                if (rc == null)
                                                    rc = row.createCell(cc + 1);
                                                if (getCellHeight(sheet, rc) != getCellHeight(sheet, cell) && cc != width - 1)
                                                    break;

                                                if (!d.equals("") && !getCellValue(sheet, rc).equals("") && !getCellValue(sheet, rc).equals(getCellValue(sheet, cell))) {
                                                    break;
                                                }
                                                if (!getCellValue(sheet, rc).equals("") && !getCellValue(sheet, cell).equals("") && !isInOneMerged(sheet, rc, cell)) {
                                                    break;
                                                }
                                                cc++;
                                                cell = row.getCell(cc);
                                                if (cell == null) cell = row.createCell(cc);
                                                st = cell.getCellStyle();
                                            }
                                            ce = cc;
                                            cell = row.getCell(cb);
                                            if (cell == null)
                                                cell = row.createCell(cb);
                                            re = rb + getVertSize(sheet, cell) - 1;
                                            if (cc == width && d.equals("") && ind != -1) {
                                                d = getCellValue(sheet, row.createCell(ind));
                                                cb -= getCellWidth(sheet, row.getCell(ind));
                                            }
                                            if (cc != width && d.equals(""))
                                                d = getCellValue(sheet, cell);
                                            makeNewMerge(sheet, rb, re, cb, ce, d);
                                            vert_expansion(sheet, cell, h_end);
                                            d = "";
                                        }
                                    }
                                    //Data Horizontal expansion
                                    DataFormatter formatter = new DataFormatter();
                                    FormulaEvaluator evaluator1 = wb_base.getCreationHelper().createFormulaEvaluator();
                                    HSSFRow r_up = sheet.getRow(h_end - 1);
                                    for (int rr = d_beg - 1; rr < d_end; rr++) {
                                        int cb, ce, rb = rr, re = rr;
                                        for (int cc = 0; cc < width; cc++) {
                                            row = sheet.getRow(rr);
                                            if (row == null)
                                                row = sheet.createRow(rr);
                                            cell = row.getCell(cc);
                                            if (cell == null) cell = row.createCell(cc);
                                            HSSFCellStyle st;
                                            st = cell.getCellStyle();
                                            cb = cc;
                                            ce = cc;
                                            String d = "", k = "";
                                            double g = 0;
                                            boolean fff = false;
                                            HSSFCell up = r_up.getCell(cc);
                                            if (up == null) up = r_up.createCell(cc);
                                            if (getCellWidth(sheet, up) == 1) {
                                                if (cell.getCellType() == FORMULA) {
                                                    HSSFSheet s_base = wb_base.getSheet(wb.getSheetName(s));
                                                    HSSFRow r_base = s_base.getRow(rr);
                                                    HSSFCell c_base = r_base.getCell(cc);
                                                    d = formatter.formatCellValue(c_base, evaluator1);
                                                    cell.setCellType(NUMERIC);
                                                    putCellValue(cell, d);
                                                }
                                            } else {
                                                for (int i = cc; i < cc + getCellWidth(sheet, up); i++) {
                                                    cell = row.getCell(i);
                                                    if (cell == null)
                                                        cell = row.createCell(i);
                                                    k = getCellValue(sheet, cell);

                                                    if (cell.getCellType() == NUMERIC) {
                                                        d = formatter.formatCellValue(cell);
                                                        fff = true;
                                                    }
                                                    if (cell.getCellType() == FORMULA) {
                                                        HSSFSheet s_base = wb_base.getSheet(wb.getSheetName(s));
                                                        HSSFRow r_base = s_base.getRow(rr);
                                                        HSSFCell c_base = r_base.getCell(i);
                                                        d = formatter.formatCellValue(c_base, evaluator1);

                                                        fff = true;
                                                    }
                                                }
                                                cc += getCellWidth(sheet, up) - 1;
                                                ce = cb + getCellWidth(sheet, up) - 1;
                                                if (fff) {
                                                    makeNewMerge(sheet, rr, rr, cb, ce, d);
                                                } else {
                                                    makeNewMerge(sheet, rr, rr, cb, ce, k);
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                        }
                    }
                }
                wb.write(fileOut);
                fileOut.close();
            }
        }
    }
}