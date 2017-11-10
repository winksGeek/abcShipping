/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package abc.shippingdocs.plants;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

/**
 *
 * @author dwink
 */
public class Plant {

    private String name;
    private String abbrev;
    private ArrayList<DailyInventory> inventory;
    private HashMap<Integer, Date> lots;
    private int mostRecentLotNumber;

    public Plant(String plantName, String plantAbbrev) {
        System.out.println(plantName);
        if ("Los Banos".equals(plantName)) {
            System.out.println("here");
        }
        name = plantName;
        inventory = new ArrayList<>();
        abbrev = plantAbbrev;
        lots = new HashMap<>();
        mostRecentLotNumber = 0;
    }

    public void processInventory(Sheet sheet, int startRow, int startColumn, FormulaEvaluator eval) {
        for (int i = startRow; i < sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            inventory.add(new DailyInventory(row, startColumn, eval));
        }
        buildInventoryMap();
    }

    public void buildInventoryMap() {
        for (DailyInventory di : inventory) {
            if (di.getBeginLotNumber() > 0) {
                for (int j = di.getBeginLotNumber(); j <= di.getEndLotNumber(); j++) {
                    lots.put(j, di.getDateCollected());
                }
            }
            if(di.getEndLotNumber() > mostRecentLotNumber){
                mostRecentLotNumber = di.getEndLotNumber();
            }
        }
    }

    private class DailyInventory {

        SimpleDateFormat sdf = new SimpleDateFormat("dd-MMM-yyyy");
        private Date dateCollected;
        private int beginLotNumber;
        private int endLotNumber;
        private int lotCount;

        public DailyInventory(Row row, int startColumn, FormulaEvaluator eval) {
            Cell dateCell = row.getCell(0);
            if (!"".equals(dateCell.toString())) {
                try {
                    dateCollected = sdf.parse(dateCell.toString());
                } catch (ParseException ex) {
                    Logger.getLogger(Plant.class.getName()).log(Level.SEVERE, null, ex);
                }
                Cell beginCell = row.getCell(startColumn + 2);
                CellValue beginCellVal = eval.evaluate(beginCell);
                Cell endCell = row.getCell(startColumn + 4);
                CellValue endCellVal = eval.evaluate(endCell);
                if (beginCellVal != null && endCellVal != null) {
                    beginLotNumber = (int) beginCellVal.getNumberValue();
                    endLotNumber = (int) endCellVal.getNumberValue();
                    lotCount = endLotNumber - beginLotNumber + 1;
                    if (beginLotNumber <= 0 && endLotNumber <= 0) {
                        lotCount = 0;
                    }               
                }
            }
        }

        /**
         * @return the dateCollected
         */
        public Date getDateCollected() {
            return dateCollected;
        }

        /**
         * @return the beginLotNumber
         */
        public int getBeginLotNumber() {
            return beginLotNumber;
        }

        /**
         * @return the endLotNumber
         */
        public int getEndLotNumber() {
            return endLotNumber;
        }

        /**
         * @return the lotCount
         */
        public int getLotCount() {
            return lotCount;
        }
    }

}
