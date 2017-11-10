/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package abc.shippingdocs.utilities;

import abc.shippingdocs.plants.Plant;
import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

/**
 *
 * @author dwink
 */
public class DataManager {

    private Workbook workbook;
    private ArrayList<Plant> plants;
    private FormulaEvaluator eval;

    private DataManager() {
        plants = new ArrayList<>();
    }

    private static class DataManagerHolder {

        private static final DataManager INSTANCE = new DataManager();
    }

    public static DataManager getInstance() {
        return DataManagerHolder.INSTANCE;
    }

    public void loadWorkbook(File file) throws FileNotFoundException, IOException {
        try {
            this.workbook = WorkbookFactory.create(file);
            eval = this.workbook.getCreationHelper().createFormulaEvaluator();
            eval.setIgnoreMissingWorkbooks(true);
            processPlants();
            System.out.println("here");
        } catch (InvalidFormatException ex) {
            Logger.getLogger(DataManager.class.getName()).log(Level.SEVERE, null, ex);
        } catch (EncryptedDocumentException ex) {
            Logger.getLogger(DataManager.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    private void processPlantInventory() {
        Sheet monday = this.workbook.getSheetAt(0);
        Sheet tuesday = this.workbook.getSheetAt(1);
        int mondayNameRow = 5;
        int tuesdayNameRow = 1;
        ArrayList<String> plantNames = new ArrayList<>();
        populatePlantNames(monday, mondayNameRow, plantNames);
        populatePlantNames(tuesday, tuesdayNameRow, plantNames);
        
    }
    
    private void processPlants(){        
        Sheet monday = this.workbook.getSheetAt(0);
        Sheet tuesday = this.workbook.getSheetAt(1);
        int mondayNameRow = 5;
        int tuesdayNameRow = 1;
        populatePlantInventory(monday, mondayNameRow, plants);
        populatePlantInventory(tuesday, tuesdayNameRow, plants);
    }

    private void populatePlantNames(Sheet sheet, int rowIndex, ArrayList<String> plantNames) {
        Row titleRow = sheet.getRow(rowIndex);
        Iterator<Cell> cIt = titleRow.cellIterator();
        while(cIt.hasNext()){
            Cell c = cIt.next();
            if(!"".equals(c.toString())){
                plantNames.add(c.toString());
            }
        }
    }

    private void populatePlantInventory(Sheet sheet, int rowIndex, ArrayList<Plant> plants) {
        Row titleRow = sheet.getRow(rowIndex);
        Iterator<Cell> cIt = titleRow.cellIterator();
        while(cIt.hasNext()){
            Cell c = cIt.next();
            if(!"".equals(c.toString())){
                Plant plant = new Plant(c.toString(), "test");
                plant.processInventory(sheet, rowIndex+4, c.getColumnIndex(), eval);
                plants.add(plant);
            }
        }
    }

}
