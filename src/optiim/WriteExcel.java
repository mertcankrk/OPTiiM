/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package optiim;

import java.io.File;
import java.io.IOException;
import java.util.Locale;

import jxl.CellView;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.UnderlineStyle;
import jxl.write.Formula;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

/**
 *
 * @author Mert
 */
public class WriteExcel {

    private WritableCellFormat timesBoldUnderline;
    private WritableCellFormat times;
    private String inputFile;

    public void setOutputFile(String inputFile) {
        this.inputFile = inputFile;
    }

    public void write(Person[] P) throws IOException, WriteException {
        File file = new File(inputFile);
        WorkbookSettings wbSettings = new WorkbookSettings();

        wbSettings.setLocale(new Locale("en", "EN"));

        WritableWorkbook workbook = Workbook.createWorkbook(file, wbSettings);
        workbook.createSheet("Report", 0);
        WritableSheet excelSheet = workbook.getSheet(0);
        createLabel(excelSheet);
        createContent(excelSheet, P);

        workbook.write();
        workbook.close();
    }

    private void createLabel(WritableSheet sheet)
            throws WriteException {

        WritableFont times10pt = new WritableFont(WritableFont.TIMES, 10);

        times = new WritableCellFormat(times10pt);

        times.setWrap(true);

        WritableFont times10ptBoldUnderline = new WritableFont(
                WritableFont.TIMES, 10, WritableFont.BOLD, false,
                UnderlineStyle.SINGLE);
        timesBoldUnderline = new WritableCellFormat(times10ptBoldUnderline);

        timesBoldUnderline.setWrap(true);

        CellView cv = new CellView();
        cv.setFormat(times);
        cv.setFormat(timesBoldUnderline);
        cv.setAutosize(true);

        addCaption(sheet, 0, 0, "AD");
        addCaption(sheet, 1, 0, "SOYAD");
        addCaption(sheet, 2, 0, "DOGUM TARIHI");
        addCaption(sheet, 3, 0, "DOGUM YERI");
        addCaption(sheet, 4, 0, "MAIL");
        addCaption(sheet, 5, 0, "TELEFON");
        addCaption(sheet, 6, 0, "DURUM");
        addCaption(sheet, 7, 0, "CALISMA DURUMU");
        addCaption(sheet, 8, 0, "UNIVERSITE");

    }

    private void createContent(WritableSheet sheet, Person[] P) throws WriteException,
            RowsExceededException {
    
        for (int i = 0; i < P.length; i++) {
            addLabel(sheet, 0, i + 1, P[i].Ad);
            addLabel(sheet, 1, i + 1, P[i].Soyad);
            addLabel(sheet, 2, i + 1, P[i].Dogum_tarihi);
            addLabel(sheet, 3, i + 1, P[i].Dogum_yeri);
            addLabel(sheet, 4, i + 1, P[i].Mail);
            addLabel(sheet, 5, i + 1, P[i].Telefon);
            addLabel(sheet, 6, i + 1, P[i].Durum);
            addLabel(sheet, 7, i + 1, P[i].Calisma_durumu);
            addLabel(sheet, 8, i + 1, P[i].Universite);
        }

     
    }

    private void addCaption(WritableSheet sheet, int column, int row, String s)
            throws RowsExceededException, WriteException {
        Label label;
        label = new Label(column, row, s, timesBoldUnderline);
        sheet.addCell(label);
    }

    private void addNumber(WritableSheet sheet, int column, int row,
            Integer integer) throws WriteException, RowsExceededException {
        Number number;
        number = new Number(column, row, integer, times);
        sheet.addCell(number);
    }

    private void addLabel(WritableSheet sheet, int column, int row, String s)
            throws WriteException, RowsExceededException {
        Label label;
        label = new Label(column, row, s, times);
        sheet.addCell(label);
    }

}
