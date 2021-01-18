
package com.lgs.workers;

import com.lgs.utils.DOMUtils;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.InputStream;
import javax.swing.JButton;
import javax.swing.JTextArea;
import javax.swing.SwingWorker;
import javax.xml.parsers.DocumentBuilderFactory;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Entities;

/**
 *
 * @author ShivanshuJS
 */
public class ConversionWorker extends SwingWorker<Void, Void>{
    
    // Swing UI components which are required to access throughout the class.
    private final String selectedFilePath;
    private final JTextArea infoTextArea;
    private final JButton fileSelectorButton;
    
    private final DOMUtils domUtils = DOMUtils.getInstance();
    
    // Constructor to initialize the global objects.
    public ConversionWorker(String selectedFilePath, JTextArea infoTextArea, JButton fileSelectorButton){
        this.selectedFilePath = selectedFilePath;
        this.infoTextArea = infoTextArea;
        this.fileSelectorButton = fileSelectorButton;
        this.setProgress(10);
    }
    
    /**
     * This method uses Apache POI library to convert Excel File to HTML File.The actual conversion is done in this method.
     */
    @Override
    protected Void doInBackground() throws Exception {
        ExcelToHtmlConverterImpl excelToHtmlConverter = new ExcelToHtmlConverterImpl(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
        String processedExcelFilePath = excelToHtmlConverter.processExcelFile(this.selectedFilePath);
        this.setProgress(20);
        InputStream inputStream = new FileInputStream(processedExcelFilePath);
        HSSFWorkbook excelWorkBook = new HSSFWorkbook(inputStream);
        excelToHtmlConverter.processSheet(excelWorkBook.getSheetAt(0));
        this.setProgress(30);
        excelToHtmlConverter.extractImages(excelWorkBook, processedExcelFilePath.substring(processedExcelFilePath.lastIndexOf("_") + 1, processedExcelFilePath.lastIndexOf(".")));
        excelToHtmlConverter.filterDocument(excelToHtmlConverter.getDocument());
        this.setProgress(40);
        ByteArrayOutputStream documentStream = excelToHtmlConverter.getDocumentStream(excelToHtmlConverter.getDocument());
        String htmlFilePath = excelToHtmlConverter.writeStreamToFile(documentStream, processedExcelFilePath.substring(processedExcelFilePath.lastIndexOf("\\") + 1, processedExcelFilePath.lastIndexOf(".")) + ".html");
        
        File htmlFile = new File(htmlFilePath);
        Document jsoupHtmlDocument = Jsoup.parse(htmlFile, "ISO-8859-1", "");
        jsoupHtmlDocument.outputSettings().escapeMode(Entities.EscapeMode.xhtml);
        
        jsoupHtmlDocument.select("th.rownumber").remove();
        Element mainTable = jsoupHtmlDocument.select("table.t1").get(0);
        excelToHtmlConverter.beautyfyTable(mainTable);
        this.setProgress(50);
        
        // get the liveDocName from project column and delete that column
        String liveDocName = mainTable.select("tr").get(1).child(11).text();
        liveDocName = liveDocName.substring(0, liveDocName.indexOf("(") - 1);
        this.domUtils.deleteColumn(mainTable, 11);
        this.setProgress(60);
        
        // remove the last row which contains the total number of rows
        mainTable.children().first().children().last().remove();
        
        // add the new required columns
        this.domUtils.addColumn(mainTable, "spaceName", "IOGPphase3", 0);
        this.domUtils.addColumn(mainTable, "liveDocName", liveDocName, 1);
        this.domUtils.addColumn(mainTable, "polarionWi", "", 2);
        this.domUtils.addColumn(mainTable, "reqNumbered", "no", 10);
        this.domUtils.addColumn(mainTable, "informStyle", "Text", 11);
        this.domUtils.addColumn(mainTable, "dem1", "no", 12);
        this.domUtils.addColumn(mainTable, "pageOrientation", "Portrait", 13);
        this.domUtils.addColumn(mainTable, "outLinks", "", 23);
        this.setProgress(70);
        
        // logics to implement the migration of jama polarion mapping
        mainTable.children().first().children().forEach(currentTr -> {
            if(currentTr.childrenSize() >= 6){
                boolean isAmendmentAdded = true;
                String clauseType = currentTr.child(5).text().trim();
                if(currentTr.child(8).text().trim().length() == 0){
                    if(currentTr.child(21).text().trim().length() == 0){
                        currentTr.child(8).text(currentTr.child(7).text());
                    } else {
                        currentTr.child(8).text(currentTr.child(21).text());
                    }
                }
                if(null != clauseType)switch (clauseType) {
                    case "Technical Requirement":
                        currentTr.child(2).text("requirement");
                        if("".equals(currentTr.child(3).text().trim()) && "".equals(currentTr.child(6).text().trim())){
                            isAmendmentAdded = false;
                        } else if("".equals(currentTr.child(3).text().trim())){
                            currentTr.before(this.domUtils.getNewRow(currentTr.child(0).text(), currentTr.child(1).text(), "amendment", currentTr.child(4).text(), clauseType, currentTr.child(6).text(), "", currentTr.child(18).text(), "", currentTr.child(20).text()));
                        } else {
                            currentTr.before(this.domUtils.getNewRow(currentTr.child(0).text(), currentTr.child(1).text(), "amendment", currentTr.child(4).text(), clauseType, currentTr.child(3).text(), "", currentTr.child(18).text(), "", currentTr.child(20).text()));
                        }
                        if(!"".equals(currentTr.child(19).text())){
                            currentTr.after(this.domUtils.getNewRow(currentTr.child(0).text(), currentTr.child(1).text(), "verification", currentTr.child(4).text(), clauseType, currentTr.child(17).text(), "", currentTr.child(18).text(), currentTr.child(19).text(), ""));
                        }
                        if(!"".equals(currentTr.child(15).text())){
                            currentTr.after(this.domUtils.getNewRow(currentTr.child(0).text(), currentTr.child(1).text(), "informative", currentTr.child(4).text(), clauseType, "", currentTr.child(15).text(), currentTr.child(18).text(), "", ""));
                        }
                        String companionGuidelineText = currentTr.child(14).text();
                        if("".equals(currentTr.child(8).text())){
                            currentTr.child(8).html("<p><em>" + companionGuidelineText + "<em></p>");
                        } else {
                            currentTr.child(8).html("<p>" + currentTr.child(8).text() + "</p><p><em>" + companionGuidelineText + "<em></p>");
                        }   
                        currentTr.child(15).text("");
                        currentTr.child(17).text("");
                        currentTr.child(19).text("");
                        currentTr.child(20).text("");
                        if(isAmendmentAdded){
                            currentTr.child(23).text("invokes:JamaTestImport/" + currentTr.child(18).text() + "A" + ",parent:JamaTestImport/" + currentTr.child(18).text() + "A");
                        }
                        break;
                    case "Text":
                        currentTr.child(2).text("requirement");
                        break;
                    case "Set":
                        currentTr.child(2).text("heading " + currentTr.child(4).text().chars().filter(character -> character == '.').count());
                        if(!"".equals(currentTr.child(6).text().trim())){
                            currentTr.before(this.domUtils.getNewRow(currentTr.child(0).text(), currentTr.child(1).text(), "amendment", currentTr.child(4).text(), clauseType, currentTr.child(6).text(), "", currentTr.child(18).text(), "", currentTr.child(20).text()));
                        } else {
                            isAmendmentAdded = false;
                        }
                        if(isAmendmentAdded){
                            currentTr.child(23).text("invokes:JamaTestImport/" + currentTr.child(18).text() + "A" + ",parent:JamaTestImport/" + currentTr.child(18).text() + "A");
                        }
                        currentTr.child(21).html("<p>" + currentTr.child(7).text() + "</p><p>" + currentTr.child(21).text() + "</p>");
                        currentTr.child(7).text("");
                        break;
                    case "Folder":
                        currentTr.child(2).text("heading " + currentTr.child(4).text().chars().filter(character -> character == '.').count());
                        if("".equals(currentTr.child(3).text().trim()) && "".equals(currentTr.child(6).text().trim())){
                            isAmendmentAdded = false;
                        } else if("".equals(currentTr.child(3).text().trim())){
                            currentTr.before(this.domUtils.getNewRow(currentTr.child(0).text(), currentTr.child(1).text(), "amendment", currentTr.child(4).text(), clauseType, currentTr.child(6).text(), "", currentTr.child(18).text(), "", currentTr.child(20).text()));
                        } else {
                            currentTr.before(this.domUtils.getNewRow(currentTr.child(0).text(), currentTr.child(1).text(), "amendment", currentTr.child(4).text(), clauseType, currentTr.child(3).text(), "", currentTr.child(18).text(), "", currentTr.child(20).text()));
                        }
                        if(isAmendmentAdded){
                            currentTr.child(23).text("invokes:JamaTestImport/" + currentTr.child(18).text() + "A" + ",parent:JamaTestImport/" + currentTr.child(18).text() + "A");
                        }
                        currentTr.child(21).html("<p>" + currentTr.child(7).text() + "</p><p>" + currentTr.child(21).text() + "</p>");
                        currentTr.child(7).text("");
                        break;
                    case "Terms and Definitions":
                        currentTr.child(2).text("definition");
                        if("".equals(currentTr.child(3).text().trim()) && "".equals(currentTr.child(6).text().trim())){
                            isAmendmentAdded = false;
                        } else if("".equals(currentTr.child(3).text().trim())){
                            currentTr.before(this.domUtils.getNewRow(currentTr.child(0).text(), currentTr.child(1).text(), "amendment", currentTr.child(4).text(), clauseType, currentTr.child(6).text(), "", currentTr.child(18).text(), "", currentTr.child(20).text()));
                        } else {
                            currentTr.before(this.domUtils.getNewRow(currentTr.child(0).text(), currentTr.child(1).text(), "amendment", currentTr.child(4).text(), clauseType, currentTr.child(3).text(), "", currentTr.child(18).text(), "", currentTr.child(20).text()));
                        }   
                        if(!"".equals(currentTr.child(15).text())){
                            currentTr.after(this.domUtils.getNewRow(currentTr.child(0).text(), currentTr.child(1).text(), "informative", currentTr.child(4).text(), clauseType, "", currentTr.child(15).text(), currentTr.child(18).text(), "", ""));
                        }
                        currentTr.child(15).text("");
                        if(isAmendmentAdded){
                            currentTr.child(23).text("invokes:JamaTestImport/" + currentTr.child(18).text() + "A" + ",parent:JamaTestImport/" + currentTr.child(18).text() + "A");
                        }
                        break;
                    case "Normative/Bibliographic Reference":
                        currentTr.child(2).text("definition");
                        break;
                    default:
                        break;
                }
            }
        });
        
        // establish parent heading outLinks
        mainTable.children().first().children().forEach(currentTr -> {
            if(currentTr.childrenSize() >= 6){
                if(currentTr.child(4).text().chars().filter(character -> character == '.').count() > 1){
                    String headingNumber = currentTr.child(4).text();
                    String currentOutLink = "";
                    if(!"".equals(currentTr.child(23).text())){
                        currentOutLink = "," + currentTr.child(23).text();
                    }
                    currentTr.child(23).text("parent:JamaTestImport/" + this.domUtils.getLegacyIdByHeadingNumber(mainTable.children().first().children(), headingNumber.substring(0, headingNumber.lastIndexOf("."))) +  currentOutLink);
                }
            }
        });
        this.setProgress(80);
        
        // update the reqNumberd column value based on reqType.
        this.domUtils.updateReqNumberColumn(mainTable.children().first().children());
        
        // add the new table header with bold and the remove the old table header
        mainTable.children().first().before("<thead><tr><th>spaceName</th><th>liveDocName</th><th>polarionWi</th><th>aboveSectionAmendment</th><th>paragraphNumber</th><th>clauseType</th><th>belowSectionAmendment</th><th>name</th><th>description</th><th>term</th><th>reqNumbered</th><th>informStyle</th><th>dem1</th><th>pageOrientation</th><th>companionGuidelineText</th><th>informDescription</th><th>requirementType</th><th>proposedVerification</th><th>legacyId</th><th>verificationMethod</th><th>action</th><th>title</th><th>sectionNumber</th><th>outLinks</th><th>tier</th></tr></thead>");
        mainTable.child(1).children().first().remove();
        this.setProgress(90);
        
        //final changes which includes removal and renaming of not required columns
        this.domUtils.deleteColumn(mainTable, 3);
        this.domUtils.deleteColumn(mainTable, 3);
        this.domUtils.deleteColumn(mainTable, 3);
        this.domUtils.deleteColumn(mainTable, 3);
        this.domUtils.deleteColumn(mainTable, 10);
        this.domUtils.deleteColumn(mainTable, 19);
        this.domUtils.renameColumn(mainTable, "polarionWi", "clauseType");
        this.domUtils.renameColumn(mainTable, "name", "paragraphNumber");
        this.domUtils.renameColumn(mainTable, "requirementType", "jamaRequirementType");
        this.domUtils.renameColumn(mainTable, "title", "jamaTitle");
        this.domUtils.renameColumn(mainTable, "action", "jamaAction");
        
        // write the DOM to a html file
        FileWriter fileWriter = new FileWriter(htmlFile);
        fileWriter.write(jsoupHtmlDocument.outerHtml());
        fileWriter.flush();
        
        // Delete the temp excel file.
        new File(processedExcelFilePath).delete();
        
        // once everything is done, update the app.
        this.infoTextArea.setText(this.infoTextArea.getText() + "\nConverted File: " + htmlFile.getAbsolutePath());
        this.fileSelectorButton.setEnabled(true);
        this.setProgress(100);
        
        return null;
    }
}
