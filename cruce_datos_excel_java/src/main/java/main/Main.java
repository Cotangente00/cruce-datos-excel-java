package main;

import model.*;

//import org.apache.poi.hssf.usermodel.HSSFWorkbook;




public class Main {
    
    public static void main(String[] args) {
        String inputFilePath = "O:/aa/test-lunes-jueves.xls";
        String outputFilePath = "O:/aa/result(xls).xlsx";
        try {
            procesamiento_hojas.boton_excel(inputFilePath, outputFilePath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}