/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package abc.shippingdocs.utilities;

import abc.shippingdocs.plants.Plant;
import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
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
    }

    private static class DataManagerHolder {

        private static final DataManager INSTANCE = new DataManager();
    }

    public static DataManager getInstance() {
        return DataManagerHolder.INSTANCE;
    }

    public void loadWorkbook(File file, String referencedFilePath) throws FileNotFoundException, IOException {
        try {
            this.workbook = WorkbookFactory.create(file);
            plants = new ArrayList<>();
            eval = this.workbook.getCreationHelper().createFormulaEvaluator();
            Map<String, FormulaEvaluator> workbooks = new HashMap();
            workbooks.put(file.getName(), eval);
            addReferencedWorkbooks(eval, referencedFilePath, workbooks);
            eval.setupReferencedWorkbooks(workbooks);
            eval.evaluateAll();
            processPlants();
        } catch (InvalidFormatException ex) {
            Logger.getLogger(DataManager.class.getName()).log(Level.SEVERE, null, ex);
        } catch (EncryptedDocumentException ex) {
            Logger.getLogger(DataManager.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public void addReferencedWorkbooks(FormulaEvaluator eval, String referencedFilesDirectory, Map<String, FormulaEvaluator> workbooks) throws IOException, InvalidFormatException {
        File directory = new File(referencedFilesDirectory);
        File[] allFiles = directory.listFiles();
        for (int i = 0; i < allFiles.length; i++) {
            File f = allFiles[i];
            if (f.isDirectory()) {
                addReferencedWorkbooks(eval, f.getCanonicalPath(), workbooks);
            } else {
                String path = f.getCanonicalPath();
                String key = createKey(path);
                File file = new File(path);
                workbooks.put(key, WorkbookFactory.create(file).getCreationHelper().createFormulaEvaluator());
                workbooks.put(key.replace("xlsm", "xlsx"), WorkbookFactory.create(file).getCreationHelper().createFormulaEvaluator());
            }
        }
    }

    public String createKey(String path) {
        String keyPath = path.replace(" ", "%20");
        String[] splitPath = keyPath.split("\\\\");
        boolean addToKey = false;
        String key = "";
        for (String part : splitPath) {
            if (addToKey || "Blood%20Reporting".equals(part)) {
                addToKey = true;
                key += "/" + part;
            }
        }
        return key;
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

    private void processPlants() {
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
        while (cIt.hasNext()) {
            Cell c = cIt.next();
            if (!"".equals(c.toString())) {
                plantNames.add(c.toString());
            }
        }
    }

    private void populatePlantInventory(Sheet sheet, int rowIndex, ArrayList<Plant> plants) {
        Row titleRow = sheet.getRow(rowIndex);
        Iterator<Cell> cIt = titleRow.cellIterator();
        while (cIt.hasNext()) {
            Cell c = cIt.next();
            if (!"".equals(c.toString())) {
                Plant plant = new Plant(c.toString(), "test");
//                plant.processInventory(sheet, rowIndex + 4, c.getColumnIndex(), eval);
                plants.add(plant);
            }
        }
    }
    
    public ArrayList<String> getPlantNames(){
        ArrayList<String> names = new ArrayList<>();
        for(Plant plant:this.plants){
            names.add(plant.getName());
        }
        return names;
    }

}
