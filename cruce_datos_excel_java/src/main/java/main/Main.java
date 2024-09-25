package main;

import funciones_boton.*;

//import org.apache.poi.hssf.usermodel.HSSFWorkbook;




public class Main {
    
    public static void main(String[] args) {
        String inputFilePath = "O:/aa/test-lunes-jueves.xlsx";
        String outputFilePath = "O:/aa/result.xlsx";
        try {
            boton.boton_excel(inputFilePath, outputFilePath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}