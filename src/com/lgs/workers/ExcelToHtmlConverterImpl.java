
package com.lgs.workers;

import com.lgs.utils.DOMUtils;
import com.lgs.utils.ExcelUtils;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintStream;
import java.util.Date;
import java.util.List;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import org.apache.poi.hssf.converter.ExcelToHtmlConverter;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.usermodel.Picture;
import org.jsoup.nodes.Element;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;

/**
 *
 * @author ShivanshuJS
 */
public class ExcelToHtmlConverterImpl extends ExcelToHtmlConverter{
    
    
    private final ExcelUtils excelUtils = ExcelUtils.getInstance();
    private final DOMUtils domUtils = DOMUtils.getInstance();
    
    // Contructor to invoke super class constructor
    public ExcelToHtmlConverterImpl(Document doc) {
        super(doc);
    }
    
    /**
     * This method is used to create a temp copy of the given excel file after removing unwanted rows.
     * @param excelFilePath The path of excel file for which the processing needs to be done.
     * @return The path of temp copy of the given excel file.
     * @throws java.io.FileNotFoundException 
     */
    public String processExcelFile(String excelFilePath) throws FileNotFoundException, IOException{
        InputStream inputStream = new FileInputStream(excelFilePath);
        HSSFWorkbook excelWorkBook = new HSSFWorkbook(inputStream);
        this.excelUtils.removeRow(excelWorkBook.getSheetAt(0), 0);
        this.excelUtils.removeRow(excelWorkBook.getSheetAt(0), 1);
        String processedExcelFilePath = System.getProperty("java.io.tmpdir") + "exceltohtml" + File.separator + excelFilePath.substring(excelFilePath.lastIndexOf("\\") + 1, excelFilePath.lastIndexOf(".")) + "_" + new Date().getTime() + ".xls";
        File processedExcelFile = new File(processedExcelFilePath);
        if(!processedExcelFile.getParentFile().exists()){
            processedExcelFile.getParentFile().mkdirs();
        }
        FileOutputStream fileOutputStream = new FileOutputStream(processedExcelFilePath);
        excelWorkBook.write(fileOutputStream);
        fileOutputStream.close();
        return processedExcelFilePath;
    }
    
    /**
     * This overridden method is used to just convert to HTML only the given sheet.
     * @param sheetToProcess The sheet to convert.
     */
    @Override
    public void processSheet(HSSFSheet sheetToProcess){
        super.processSheet(sheetToProcess);
    }
    
    /**
     * This method is used to extract all of the images inserted in the excel file.
     * @param excelWorkBook The excel workbook from which the extraction is to be done.
     * @param directoryIdentifier Unique identifier to understand which directory belongs to which converted HTML file.
     * @throws java.io.FileNotFoundException
     */
    public void extractImages(HSSFWorkbook excelWorkBook, String directoryIdentifier) throws FileNotFoundException, IOException{
        String imageFilePath = System.getProperty("java.io.tmpdir") + "exceltohtml" + File.separator + "images_" + directoryIdentifier + File.separator;
        List pictures = excelWorkBook.getAllPictures();
        if(pictures != null && pictures.size() > 0){
            for(int i=0; i<pictures.size(); i++){
                Picture picture = (Picture) pictures.get(i);
                File imageFile = new File(imageFilePath + picture.suggestFullFileName());
                if(!imageFile.getParentFile().exists()){
                    imageFile.getParentFile().mkdirs();
                }
                picture.writeImageContent(new FileOutputStream(imageFile));
            }
        }
    }
    
    /**
     * This method is used to filter out all of the document headings and col groups.
     * @param htmlDocument The HTML document for which the filter is to be done.
     */
    public void filterDocument(Document htmlDocument){
        // Remove Sheet Heading
        NodeList nodeList = htmlDocument.getElementsByTagName("h2");
        for(int i=0; i<nodeList.getLength(); i++){
            if(nodeList.item(i).getTextContent().contains("Sheet")){
                nodeList.item(i).getParentNode().removeChild(nodeList.item(i));
            }
        }
        // Remove colgroup tag
        nodeList = htmlDocument.getElementsByTagName("colgroup");
        for(int i=0; i<nodeList.getLength(); i++){
            nodeList.item(i).getParentNode().removeChild(nodeList.item(i));
        }
    }
    
    /**
     * This method converts a given HTML document to to Stream.
     * @param htmlDocument The HTML document for which the conversion is to be done.
     * @return The converted ByteArrayOutputStream
     * @throws javax.xml.transform.TransformerConfigurationException
     * @throws java.io.IOException
     * @throws javax.xml.transform.TransformerException
     */
    public ByteArrayOutputStream getDocumentStream(Document htmlDocument) throws TransformerConfigurationException, TransformerException, IOException{
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        DOMSource domSource = new DOMSource(htmlDocument);
        StreamResult streamResult = new StreamResult(byteArrayOutputStream);
        TransformerFactory transformerFactory = TransformerFactory.newInstance();
        Transformer serializer = transformerFactory.newTransformer();
        serializer.setOutputProperty(OutputKeys.ENCODING, "ISO-8859-1");
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");
        serializer.transform(domSource, streamResult);
        byteArrayOutputStream.close();
        return byteArrayOutputStream;
    }
    
    /**
     * This method is used to write an array of Stream to a file.
     * @param documentStream The Stream which needs to be written to file.
     * @param fileName name of the file.
     * @return Path of the new file created.
     * @throws java.io.FileNotFoundException
     */
    public String writeStreamToFile(ByteArrayOutputStream documentStream, String fileName) throws FileNotFoundException{
        File htmlFile = new File(System.getProperty("java.io.tmpdir") + "exceltohtml" + File.separator + fileName);
        OutputStream outputStream = new FileOutputStream(htmlFile);
        PrintStream printStream = new PrintStream(outputStream);
        printStream.print(new String(documentStream.toByteArray()));
        printStream.close();
        return htmlFile.getAbsolutePath();
    }
    
    /**
     * This method is used to apply some styles to the main table.
     * @param tableElement The table object on which styles to be applied.
     */
    public void beautyfyTable(Element tableElement){
        // beautyfy the table table
        tableElement.attr("border", "1");
        tableElement.attr("cellspacing", "0");
        tableElement.attr("cellpadding", "5");
        
        // remove the unneccessary columns
        String[] columnsToDelete = new String[]{"B", "E", "F", "I", "J", "N", "P", "Q", "R", "T", "W", "X", "Y", "Z", "AB", "AC", "AD", "AE", "AG", "AJ", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS"};
        for(String columnToDelete: columnsToDelete){
            this.domUtils.deleteColumn(tableElement, columnToDelete);
        }
        
        // 
        tableElement.children().first().remove();
        tableElement.children().first().children().first().remove();
        
        // rename the columns as per the name map
        this.domUtils.getColumnNameMap().forEach((key, value) -> {
            this.domUtils.renameColumn(tableElement, key, value);
        });
    }
}