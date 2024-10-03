/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.cruce_datos_excel.procesamiento_hojas.maquillaje;

import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.*;

/**
 *
 * @author jcavilaa
 */
public class order_Hoja1 {
    public static void reorganizeExcel_Hoja1(Workbook wb) throws IOException {
        
        Sheet originalSheet = wb.getSheetAt(1);  // Obtener la primera hoja
        Sheet newSheet = wb.createSheet("ReorganizedSheet");  // Crear una nueva hoja para los datos reorganizados

        List<Row> rowsWithEmptyH = new ArrayList<>();
        List<Row> rowsWithData = new ArrayList<>();
        List<Row> data = new ArrayList<>();
        

        // Iterar sobre la columna D para determinar el rango de filas
        int rowIndex = 4;  // Empezar desde la fila 5 (índice 4)
        while (true) {
            Row row = originalSheet.getRow(rowIndex);
            if (row == null || row.getCell(3) == null || row.getCell(3).getCellType() == Cell.CELL_TYPE_BLANK) {
                break;  // Detener cuando se encuentre la primera celda vacía en la columna D
            }

            // Verificar las celdas en la columna H (índice 7 respectivamente)
            Cell cellH = row.getCell(7);

            boolean isCellHEmpty = (cellH == null || cellH.getCellType() == Cell.CELL_TYPE_BLANK || cellH.getStringCellValue().isEmpty());

            // Si las celdas están vacías, agregar la fila a la lista correspondiente
            if (isCellHEmpty) {
                rowsWithEmptyH.add(row);
            } else {
                rowsWithData.add(row);
            }   

            rowIndex++;
        }


        int newRowIndex = 4; 
        // Copiar filas con celdas vacías en H al principio
        for (Row row : rowsWithEmptyH) {
            order_INFORME_SOLICITUDES.copyRow(row, newSheet.createRow(newRowIndex++), wb);
        }

        // Copiar filas con datos después
        for (Row row : rowsWithData) {
            order_INFORME_SOLICITUDES.copyRow(row, newSheet.createRow(newRowIndex++), wb);
        }




        // Obtener la última fila y columna de la hoja
        int lastRow = originalSheet.getLastRowNum();
        // Iterar sobre las filas y eliminar todas las celdas de la hoja original (Hoja1)
        for (int i = lastRow; i >= 0; i--) {
            Row row = originalSheet.getRow(i);
            if (row != null) {
                originalSheet.removeRow(row);
            }
        }

        // Almacenar todos los datos de la columna
        int rowIndex2 = 4;  // Comenzar desde la fila 4 en la hoja nueva
        while (true) {
            Row row = newSheet.getRow(rowIndex2);
            if (row == null || row.getCell(3) == null || row.getCell(3).getCellType() == Cell.CELL_TYPE_BLANK) {
                break;  // Detener cuando se encuentre la primera celda vacía en la columna A
            }
            data.add(row);
            rowIndex2++;
        }
        
        int newRowIndex2 = 4;  // Comenzar desde la fila 5 en la nueva hoja
        for (Row row : data) {
            order_INFORME_SOLICITUDES.copyRow(row, originalSheet.createRow(newRowIndex2++), wb);
        }

        wb.removeSheetAt(2);
    }
    
    public static void main(String[] args) throws Exception {
        String inputFilePath = "O:/aa/result2.xlsx";
        String outputFilePath = "O:/aa/result3.xlsx";
        Workbook wb;
        try (FileInputStream fis = new FileInputStream(new File(inputFilePath))) {
            wb = WorkbookFactory.create(fis);  // Apache POI detecta automáticamente si es .xls o .xlsx
        }

        try {
            reorganizeExcel_Hoja1(wb);
            wb.write(new FileOutputStream(outputFilePath));
            wb.close();
            System.out.println("Archivo procesado exitosamente.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}