
package com.lgs.utils;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 *
 * @author ShivanshuJS
 */
public class ExcelUtils {
    
    private static ExcelUtils onlyInstance = null;
    
    private ExcelUtils(){}
    
    /**
     * This static method can be used to get an instance of this singleton class.
     * @return an instance of this singleton class.
     */
    public static ExcelUtils getInstance(){
        if(onlyInstance == null){
            onlyInstance = new ExcelUtils();
        }
        return onlyInstance;
    }
    
    /**
     * This method is used to remove/delete a row using the row index from a given excel sheet.
     * @param sheet from which the row has to be deleted.
     * @param rowIndex the index of the row to delete.
     */ 
    public void removeRow(Sheet sheet, int rowIndex) {
        int lastRowNum = sheet.getLastRowNum();
        if(rowIndex >=0 && rowIndex < lastRowNum){
            sheet.shiftRows(rowIndex+1,lastRowNum, -1);
        }
        if(rowIndex == lastRowNum){
            HSSFRow removingRow=(HSSFRow) sheet.getRow(rowIndex);
            if(removingRow != null){
                sheet.removeRow(removingRow);
            }
        }
    }
    
    /**
     * Takes an existing Cell of an excel sheet and merges all the styles and formula into the new one.
     * @param newCell the new cell.
     * @param oldCell the existing cell.
     */
    public void cloneCell(Cell newCell, Cell oldCell){
        newCell.setCellComment(oldCell.getCellComment() );
        newCell.setCellStyle(oldCell.getCellStyle() );
        switch (newCell.getCellType()){
            case BOOLEAN:{
                newCell.setCellValue(oldCell.getBooleanCellValue() );
                break;
            }
            case NUMERIC:{
                newCell.setCellValue(oldCell.getNumericCellValue() );
                break;
            }
            case STRING:{
                newCell.setCellValue(oldCell.getStringCellValue() );
                break;
            }
            case ERROR:{
                newCell.setCellValue(oldCell.getErrorCellValue() );
                break;
            }
            case FORMULA:{
                newCell.setCellFormula(oldCell.getCellFormula() );
                break;
            } 
        }
    }
    
    /**
     * This method is used to delete a column from excel sheet.
     * @param sheet the sheet from which the column has to be removed.
     * @param columnIndex index of the column which has to be removed.
     */
    public void deleteColumn(Sheet sheet, int columnIndex){
        int maxColumn = 0;
        for(int r=0; r<sheet.getLastRowNum()+1; r++){
            Row row = sheet.getRow(r);
            if(row == null){
                continue;
            }
            int lastColumn = row.getLastCellNum();
            if(lastColumn > maxColumn){
                maxColumn = lastColumn;
            }
            if(lastColumn < columnIndex){
                continue;
            }
            for(int x=columnIndex+1; x<lastColumn+1; x++){
                Cell oldCell = row.getCell(x-1);
                if(oldCell != null){
                    row.removeCell( oldCell );
                }
                Cell nextCell = row.getCell(x);
                if(nextCell != null){
                    Cell newCell = row.createCell(x-1, nextCell.getCellType());
                    cloneCell(newCell, nextCell);
                }
            }
        }
        for(int c=0; c<maxColumn; c++){
            sheet.setColumnWidth(c, sheet.getColumnWidth(c+1));
        }
    }
}
