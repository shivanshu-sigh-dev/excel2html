
package com.lgs.utils;

import java.util.HashMap;
import java.util.concurrent.atomic.AtomicBoolean;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

/**
 *
 * @author ShivanshuJS
 */
public class DOMUtils {
    
    private static DOMUtils onlyInstance = null;
    
    private DOMUtils(){}
    
    /**
     * This static method can be used to get an instance of this singleton class.
     * @return an instance of this singleton class.
     */
    public static DOMUtils getInstance(){
        if(onlyInstance == null){
            onlyInstance = new DOMUtils();
        }
        return onlyInstance;
    }
    
    /**
     * This method can be used to get the new name map of the columns.
     * @return a hash-map which has key as the old name and value as the new name.
     */
    public HashMap<String, String> getColumnNameMap(){
        HashMap<String, String> nameMap = new HashMap<>();
        nameMap.put("Description", "description");
        nameMap.put("Heading", "paragraphNumber");
        nameMap.put("Item Type", "clauseType");
        nameMap.put("Below section amendment", "belowSectionAmendment");
        nameMap.put("Above section amendment", "aboveSectionAmendment");
        nameMap.put("Justification", "informDescription");
        nameMap.put("Global ID", "legacyId");
        nameMap.put("Title", "title");
        nameMap.put("Name", "name");
        nameMap.put("Term", "term");
        nameMap.put("Proposed Verification", "proposedVerification");
        nameMap.put("Verification Method", "verificationMethod");
        nameMap.put("Action", "action");
        nameMap.put("Section Number", "sectionNumber");
        nameMap.put("Tier", "tier");
        nameMap.put("Companion Guideline text", "companionGuidelineText");
        nameMap.put("Requirement type", "requirementType");
        return nameMap;
    }
    
    /**
     * This method is used to get column index of a HTML table.
     * @param firstTr the first table row which consist the column id.
     * @param columnId the columnId for which the index is required.
     * @return index of the given column id.
     */
    public int getColumnIndexById(Element firstTr, String columnId){
        int columnIndex = 0;
        Elements allTd = firstTr.children();
        for(Element td: allTd){
            if(td.text().trim().equals(columnId)){
                break;
            }
            columnIndex++;
        }
        return columnIndex;
    }
    
    /**
     * This method is used to remove/delete a HTML table column using column index.
     * @param tableElement The HTML table from which the column to be removed.
     * @param columnIndex The column index which need to be removed.
     */
    public void deleteColumn(Element tableElement, int columnIndex){
        tableElement.select("tr").forEach((row) -> {
            if(columnIndex < row.childrenSize()){
                row.child(columnIndex).remove();
            }
        });
    }
    
    /**
     * This method is used to remove/delete a HTML table column using column id.
     * @param tableElement The HTML table from which the column to be removed.
     * @param columnId The column id which need to be removed.
     */
    public void deleteColumn(Element tableElement, String columnId){
        int columnIndex = getColumnIndexById(tableElement.select("tr").get(0), columnId);
        deleteColumn(tableElement, columnIndex);
    }
    
    /**
     * This method is used to change column name of a HTML table.
     * @param tableElement The HTML table from which the column to be removed.
     * @param oldName the old name of the column.
     * @param newName the new name to set.
     */
    public void renameColumn(Element tableElement, String oldName, String newName){
        tableElement.select("tr").get(0).children().forEach((child) -> {
            if(child.text().equals(oldName)){
                child.text(newName);
            }
        });
    }
    
    /**
     * This method is used to add a column to a HTML table.
     * @param tableElement The HTML table from which the column to be removed.
     * @param newColumnName the name of the new column.
     * @param newColumnValue the value/text to be stored in that column.
     * @param beforeColumnIndex the index of column before which we have to add the new column.
     */
    public void addColumn(Element tableElement, String newColumnName, String newColumnValue, int beforeColumnIndex){
        AtomicBoolean isFirstIteration = new AtomicBoolean(true);
        tableElement.select("tr").forEach(row -> {
            if(row.childrenSize() != 0){
                if(isFirstIteration.get()){
                    row.child(beforeColumnIndex).before("<td class=\"c1\">" + newColumnName + "</td>");
                    isFirstIteration.getAndSet(false);
                } else {
                    row.child(beforeColumnIndex).before("<td class=\"c1\">" + newColumnValue + "</td>");
                }
            }
        });
    }
    
    /**
     * This method creates a new HTML table row using the values provided as parameters.
     * @param spaceName
     * @param liveDocName
     * @param polarionWi
     * @param paragraphNumber
     * @param clauseType
     * @param description
     * @param informDescription
     * @param legacyId
     * @param verificationMethod
     * @param action
     * @return the newly created row.
     */
    public String getNewRow(String spaceName, String liveDocName, String polarionWi, String paragraphNumber, String clauseType, String description, String informDescription, String legacyId, String verificationMethod, String action){
        String updatedLegacyId = "";
        String outLinks = "";
        switch(polarionWi){
            case "amendment":
                updatedLegacyId = legacyId + "A";
                break;
            case "informative":
                updatedLegacyId = legacyId + "I";
                outLinks = "<p>informs:JamaTestImport/" + legacyId + ",parent:JamaTestImport/" + legacyId + "</p>";
            break;
            case "verification":
                updatedLegacyId = legacyId + "V";
                outLinks = "<p>verifies:JamaTestImport/" + legacyId + ",parent:JamaTestImport/" + legacyId + "</p>";
            break;
            default: 
                updatedLegacyId = legacyId;
            break;
        }
        String row = 
            "<tr class=\"r1\" bgcolor=\"#FFE4C4\">"
                + "<td class=\"c1\">" + spaceName + "</td>"
                + "<td class=\"c1\">" + liveDocName + "</td>"
                + "<td class=\"c1\">" + polarionWi + "</td>"
                + "<td class=\"c1\"></td>"
                + "<td class=\"c1\"></td>"
                + "<td class=\"c1\">" + clauseType + "</td>"
                + "<td class=\"c1\"></td>"
                + "<td class=\"c1\"></td>"        
                + "<td class=\"c1\">" + description + "</td>"
                + "<td class=\"c1\"></td>"
                + "<td class=\"c1\">no</td>"
                + "<td class=\"c1\">Text</td>"
                + "<td class=\"c1\">no</td>"
                + "<td class=\"c1\">Portrait</td>"
                + "<td class=\"c1\"></td>"
                + "<td class=\"c1\">" + informDescription + "</td>"
                + "<td class=\"c1\"></td>"
                + "<td class=\"c1\"></td>"
                + "<td class=\"c1\">" + updatedLegacyId + "</td>"
                + "<td class=\"c1\">" + verificationMethod + "</td>"
                + "<td class=\"c1\">" + action + "</td>"
                + "<td class=\"c1\"></td>"
                + "<td class=\"c1\"></td>"
                + "<td class=\"c1\">" + outLinks + "</td>"
                + "<td class=\"c1\"></td>"
            + "</tr>";
        return row;
    }
    
    /**
     * This method is used to update the reqNumberd column value based on reqType.
     * @param rows collection of all rows of the table.
     */
    public void updateReqNumberColumn(Elements rows){
        rows.forEach(row -> {
            if(row.childrenSize() >= 16){
                String requirementType = row.child(16).text();
                if("Text".equals(requirementType)){
                    row.child(10).text("yes");
                }
            }
        });
    }
    
    /**
     * This method is used to get the legacy id of a given heading number.
     * @param rows collection of all rows of the table.
     * @param headingNumber Heading number for which legacy id to be found.
     * @return The legacy id.
     */
    public String getLegacyIdByHeadingNumber(Elements rows, String headingNumber){
        String legacyId = null;
        for(int i=0; i<rows.size(); i++){
            if(headingNumber.equals(rows.get(i).child(4).text())){
                legacyId = rows.get(i).child(18).text();
                break;
            }
        }
        return legacyId;
    }
}
