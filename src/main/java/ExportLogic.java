import org.apache.poi.xwpf.usermodel.*;

import javax.swing.*;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

// SINGLETON CLASS
public class ExportLogic {

  private static ExportLogic exportLogic;

  public static ExportLogic getExportLogic() {
    if(exportLogic == null) {
      exportLogic = new ExportLogic();
    }
    return exportLogic;
  }

  // Logic for the final table formatting

  public void createDocument(String studentName, String className, String trainerName,
                             String lernfeld, String berichtNr, String year,
                             String week, String mondayText, String tuesdayText,
                             String wednesdayText, String thursdayText, String fridayText) {
    // Step 1: Creating blank document
    XWPFDocument document = new XWPFDocument();
    // Step 2: Getting path of current working directory
    // to create the pdf file in the same directory
    String path = System.getProperty("user.dir");
    path += "/Berichtsheft_" + berichtNr + "_" + studentName + ".docx";
    // Step 3: Creating a file object with the path specified
    FileOutputStream out
      = null;
    try {
      out = new FileOutputStream(new File(path));
    } catch (FileNotFoundException e) {
      e.printStackTrace();
      showErrorDialog();
    }

    // Header Table
    XWPFTable headerTable = document.createTable();
    headerTable.setWidth("100%");

    XWPFTableRow tableRowOne = headerTable.getRow(0);
    tableRowOne.createCell();
    tableRowOne.addNewTableCell();

    boldHelper(tableRowOne.getCell(0), "Name: ");
    fillCellWithText(tableRowOne.getCell(0), studentName);
    // Klasse
    boldHelper(tableRowOne.getCell(1), "Klasse: " );
    fillCellWithText(tableRowOne.getCell(1), className + " ");
    boldHelper(tableRowOne.getCell(1), "Jahr: ");
    fillCellWithText(tableRowOne.getCell(1), year);
    // Nr
    boldHelper(tableRowOne.getCell(2), "Nachweis-Nr: ");
    fillCellWithText(tableRowOne.getCell(2), berichtNr);

    XWPFTableRow tableRowTwo = headerTable.createRow();

    // Lehrer
    boldHelper(tableRowTwo.getCell(0), "Lehrer/Trainer: ");
    fillCellWithText(tableRowTwo.getCell(0), trainerName);
    // Lernfeld
    boldHelper(tableRowTwo.getCell(1), "Lernfeld: ");
    fillCellWithText(tableRowTwo.getCell(1), lernfeld);
    // Von bis
    boldHelper(tableRowTwo.getCell(2), "Woche: ");
    fillCellWithText(tableRowTwo.getCell(2), week);
    document.createParagraph();

    // Body Table
    XWPFTable bodyTable = document.createTable();
    bodyTable.setWidth("100%");

    XWPFTableRow tableRowThree = bodyTable.getRow(0);
    tableRowThree.createCell();
    tableRowThree.getCell(0).setWidth("5%");
    boldHelper(tableRowThree.getCell(0), "Mo: ");
    fillCell_reatainLinebreak(mondayText, tableRowThree.getCell(1));
    XWPFTableRow tableRowFour = bodyTable.createRow();
    boldHelper(tableRowFour.getCell(0), "Di: ");
    fillCell_reatainLinebreak(tuesdayText, tableRowFour.getCell(1));
    XWPFTableRow tableRowFive = bodyTable.createRow();
    boldHelper(tableRowFive.getCell(0), "Mi: ");
    fillCell_reatainLinebreak(wednesdayText, tableRowFive.getCell(1));
    XWPFTableRow tableRowSix = bodyTable.createRow();
    boldHelper(tableRowSix.getCell(0), "Do: ");
    fillCell_reatainLinebreak(thursdayText, tableRowSix.getCell(1));
    XWPFTableRow tableRowSeven = bodyTable.createRow();
    boldHelper(tableRowSeven.getCell(0), "Fr: ");
    fillCell_reatainLinebreak(fridayText, tableRowSeven.getCell(1));

    XWPFParagraph footerParagraph = document.createParagraph();
    XWPFRun footerRun = footerParagraph.createRun();
    footerRun.addBreak();
    footerRun.setText("_____________________________" +
      "                                                   _____________________________");
    footerRun.addBreak();
    footerRun.setText("Auszubildender" +
      "                                                                                       " +
      "Ausbildungsbetrieb");

    // Step 6: Saving changes to document
    try {
      document.write(out);
    } catch (IOException e) {
      e.printStackTrace();
      showErrorDialog();
    }

    // Step 7: Closing the connections
    assert out != null;
    try {
      out.close();
    } catch (IOException e) {
      e.printStackTrace();
      showErrorDialog();
    }
    try {
      document.close();
    } catch (IOException e) {
      e.printStackTrace();
      showErrorDialog();
    }
  }

  // Helper methods to place, manipulate and format sections of text

  private void fillCellWithText(XWPFTableCell cell, String text) {
    XWPFParagraph paragraph = cell.getParagraphs().get(0);
    XWPFRun run = paragraph.createRun();
    run.setText(text);
  }

  private void fillCell_reatainLinebreak(String string, XWPFTableCell cell) {
    String[] newlineSplits = string.split("\n");
    XWPFParagraph paragraph = cell.getParagraphs().get(0);
    XWPFRun run = paragraph.createRun();
    for(String s:newlineSplits) {
      run.setText(s);
      run.addBreak();
    }
  }

  private void boldHelper(XWPFTableCell cell, String textToBold) {
    XWPFParagraph paragraph = cell.getParagraphs().get(0);
    XWPFRun run = paragraph.createRun();
    run.setText(textToBold);
    run.setBold(true);
  }

  private void showErrorDialog() {
    JOptionPane.showMessageDialog(new JFrame(), "Fehler! Ueberpruefen Sie ihre Angaben" +
      " und versuchen Sie es erneut");
  }
}