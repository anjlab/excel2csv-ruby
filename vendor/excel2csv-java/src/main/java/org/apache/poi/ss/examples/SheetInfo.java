package org.apache.poi.ss.examples;

import java.util.ArrayList;
import java.util.List;

public class SheetInfo {

    private List<List<String>> rows;
    private int maxRowWidth = 0;

    public SheetInfo() {
        this.rows = new ArrayList<List<String>>();
    }
    
    public int getRowCount() {
        return rows.size();
    }
    
    public int getMaxRowWidth() {
        return maxRowWidth;
    }
    
    public List<String> getRow(int i) {
        return rows.get(i);
    }
    
    public void addRow(List<String> csvLine) {
        // Issue #4
        while (csvLine.size() > 0 && ToCSV.isNullOrEmpty(csvLine.get(csvLine.size() - 1))) {
            csvLine.remove(csvLine.size() - 1);
        }
        int lastCellNum = csvLine.size();
        // Make a note of the index number of the right most cell. This value
        // will later be used to ensure that the matrix of data in the CSV file
        // is square.
        if(lastCellNum > this.maxRowWidth) {
            this.maxRowWidth = lastCellNum;
        }
        rows.add(csvLine);
    }
    
}
