
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi3.hssf.usermodel.*;
import org.apache.poi3.ss.util.*;
import org.apache.poi3.xssf.usermodel.*;
import org.apache.poi3.ss.usermodel.*;

public class Convert {

    private static XSSFWorkbook workbookOld = null;
    private static HSSFWorkbook workbookNew = null;
    private static int lastColumn = 0;
    private static HashMap<Integer, HSSFCellStyle> styleMap;

    public static boolean transformXlsx2Xls(String path) {
        FileInputStream fin = null;
        OutputStream out = null;
        boolean result = true;
        styleMap = new HashMap<Integer, HSSFCellStyle>();
        try {
            File f = new File(path);
            fin = new FileInputStream(f);
            String parent_path = f.getParentFile().getAbsolutePath();
            String file_name = f.getName().split("\\.")[0];
            String separator = System.getProperty("file.separator");
            String path_out = parent_path + separator + file_name + ".xls";
            File outF = new File(path_out);
            if (outF.exists()) {
                outF.delete();
            }
            workbookOld = new XSSFWorkbook(fin);
            workbookNew = new HSSFWorkbook();
            HSSFSheet sheetNew;
            XSSFSheet sheetOld;
            Convert.workbookNew.setForceFormulaRecalculation(Convert.workbookOld.getForceFormulaRecalculation());
            Convert.workbookNew.setMissingCellPolicy(Convert.workbookOld.getMissingCellPolicy());
            for (int i = 0; i < Convert.workbookOld.getNumberOfSheets(); i++) {
                sheetOld = workbookOld.getSheetAt(i);
                sheetNew = workbookNew.createSheet(sheetOld.getSheetName());
                transformXlsx2Xls(sheetOld, sheetNew);
            }
            out = new BufferedOutputStream(new FileOutputStream(outF));
            workbookNew.write(out);
        } catch (Exception e) {
            Logger.getLogger(Convert.class.getName()).log(Level.SEVERE, null, e);
            result = false;
        } finally {
            try {
                if (out != null) {
                    out.close();
                }
                if (fin != null) {
                    fin.close();
                }
            } catch (IOException ex) {
                Logger.getLogger(Convert.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
        return result;
    }

    private static void transformXlsx2Xls(XSSFSheet sheetOld, HSSFSheet sheetNew) {
        System.out.println("transform Sheet");

        sheetNew.setDisplayFormulas(sheetOld.isDisplayFormulas());
        sheetNew.setDisplayGridlines(sheetOld.isDisplayGridlines());
        sheetNew.setDisplayGuts(sheetOld.getDisplayGuts());
        sheetNew.setDisplayRowColHeadings(sheetOld.isDisplayRowColHeadings());
        sheetNew.setDisplayZeros(sheetOld.isDisplayZeros());
        sheetNew.setFitToPage(sheetOld.getFitToPage());
        sheetNew.setForceFormulaRecalculation(sheetOld
                .getForceFormulaRecalculation());
        sheetNew.setHorizontallyCenter(sheetOld.getHorizontallyCenter());
        sheetNew.setMargin(Sheet.BottomMargin,
                sheetOld.getMargin(Sheet.BottomMargin));
        sheetNew.setMargin(Sheet.FooterMargin,
                sheetOld.getMargin(Sheet.FooterMargin));
        sheetNew.setMargin(Sheet.HeaderMargin,
                sheetOld.getMargin(Sheet.HeaderMargin));
        sheetNew.setMargin(Sheet.LeftMargin,
                sheetOld.getMargin(Sheet.LeftMargin));
        sheetNew.setMargin(Sheet.RightMargin,
                sheetOld.getMargin(Sheet.RightMargin));
        sheetNew.setMargin(Sheet.TopMargin, sheetOld.getMargin(Sheet.TopMargin));
        sheetNew.setPrintGridlines(sheetNew.isPrintGridlines());
        sheetNew.setRightToLeft(sheetNew.isRightToLeft());
        sheetNew.setRowSumsBelow(sheetNew.getRowSumsBelow());
        sheetNew.setRowSumsRight(sheetNew.getRowSumsRight());
        sheetNew.setVerticallyCenter(sheetOld.getVerticallyCenter());

        HSSFRow rowNew;
        for (Row row : sheetOld) {
            rowNew = sheetNew.createRow(row.getRowNum());
            if (rowNew != null) {
                transformXlsx2Xls((XSSFRow) row, rowNew);
            }
        }

        for (int i = 0; i < lastColumn; i++) {
            sheetNew.setColumnWidth(i, sheetOld.getColumnWidth(i));
            sheetNew.setColumnHidden(i, sheetOld.isColumnHidden(i));
        }

        for (int i = 0; i < sheetOld.getNumMergedRegions(); i++) {
            CellRangeAddress merged = sheetOld.getMergedRegion(i);
            sheetNew.addMergedRegion(merged);
        }
    }

    private static void transformXlsx2Xls(XSSFRow rowOld, HSSFRow rowNew) {
        HSSFCell cellNew;
        rowNew.setHeight(rowOld.getHeight());
        if (rowOld.getRowStyle() != null) {
            Integer hash = rowOld.getRowStyle().hashCode();
            if (!styleMap.containsKey(hash)) {
                transformXlsx2Xls(hash, rowOld.getRowStyle(),
                        workbookNew.createCellStyle());
            }
            rowNew.setRowStyle(styleMap.get(hash));
        }
        for (Cell cell : rowOld) {
            cellNew = rowNew.createCell(cell.getColumnIndex(),
                    cell.getCellType());
            if (cellNew != null) {
                transformXlsx2Xls((XSSFCell) cell, cellNew);
            }
        }
        lastColumn = Math.max(lastColumn, rowOld.getLastCellNum());
    }

    private static void transformXlsx2Xls(XSSFCell cellOld, HSSFCell cellNew) {
        cellNew.setCellComment(cellOld.getCellComment());

        Integer hash = cellOld.getCellStyle().hashCode();
        if (!styleMap.containsKey(hash)) {
            transformXlsx2Xls(hash, cellOld.getCellStyle(),
                    workbookNew.createCellStyle());
        }
        cellNew.setCellStyle(styleMap.get(hash));

        switch (cellOld.getCellType()) {
            case Cell.CELL_TYPE_BLANK:
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                cellNew.setCellValue(cellOld.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_ERROR:
                cellNew.setCellValue(cellOld.getErrorCellValue());
                break;
            case Cell.CELL_TYPE_FORMULA:
                cellNew.setCellValue(cellOld.getCellFormula());
                break;
            case Cell.CELL_TYPE_NUMERIC:
                cellNew.setCellValue(cellOld.getNumericCellValue());
                break;
            case Cell.CELL_TYPE_STRING:
                cellNew.setCellValue(cellOld.getStringCellValue());
                break;
            default:
                System.out.println("");
        }
    }

    private static void transformXlsx2Xls(Integer hash, XSSFCellStyle styleOld, HSSFCellStyle styleNew) {
        styleNew.setAlignment(styleOld.getAlignment());
        styleNew.setBorderBottom(styleOld.getBorderBottom());
        styleNew.setBorderLeft(styleOld.getBorderLeft());
        styleNew.setBorderRight(styleOld.getBorderRight());
        styleNew.setBorderTop(styleOld.getBorderTop());
        styleNew.setDataFormat(Convert.transformXlsx2Xls(styleOld.getDataFormat()));
        styleNew.setFillBackgroundColor(styleOld.getFillBackgroundColor());
        styleNew.setFillForegroundColor(styleOld.getFillForegroundColor());
        styleNew.setFillPattern(styleOld.getFillPattern());
        styleNew.setFont(Convert.transformXlsx2Xls(styleOld.getFont()));
        styleNew.setHidden(styleOld.getHidden());
        styleNew.setIndention(styleOld.getIndention());
        styleNew.setLocked(styleOld.getLocked());
        styleNew.setVerticalAlignment(styleOld.getVerticalAlignment());
        styleNew.setWrapText(styleOld.getWrapText());
        Convert.styleMap.put(hash, styleNew);
    }

    private static short transformXlsx2Xls(short index) {
        DataFormat formatOld = (DataFormat) workbookOld.createDataFormat();
        DataFormat formatNew = workbookNew.createDataFormat();
        return formatNew.getFormat(formatOld.getFormat(index));
    }

    private static HSSFFont transformXlsx2Xls(XSSFFont fontOld) {
        HSSFFont fontNew = workbookNew.createFont();
        fontNew.setBoldweight(fontOld.getBoldweight());
        fontNew.setCharSet(fontOld.getCharSet());
        fontNew.setColor(fontOld.getColor());
        fontNew.setFontName(fontOld.getFontName());
        fontNew.setFontHeight(fontOld.getFontHeight());
        fontNew.setItalic(fontOld.getItalic());
        fontNew.setStrikeout(fontOld.getStrikeout());
        fontNew.setTypeOffset(fontOld.getTypeOffset());
        fontNew.setUnderline(fontOld.getUnderline());
        return fontNew;
    }
}