/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.cruce_datos_excel.procesamiento_hojas.manipular_Hoja1;

//import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;

/*
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
*/
import java.util.*;

/**
 *
 * @author jcavilaa
 */
public class buscarV_nombres {
    public static void simulateBUSCARVHoja1(Workbook wb) throws Exception {
        Sheet ws = wb.getSheetAt(1); // Hoja1
        Sheet ws2 = wb.getSheetAt(0); // INFORME SOLICITUDES

        // Obtener las columnas de interés como iteradores
        Iterator<Row> rowIterator1 = ws.iterator();
        Iterator<Row> rowIterator2 = ws2.iterator();
        rowIterator2.next(); // Saltar el encabezado de INFORME SOLICITUDES

        // Crear conjuntos para almacenar los números de documento
        Set<String> numerosHoja1 = new HashSet<>();
        Set<Map<String, String>> datosINFORME_SOLICITUDES = new HashSet<>();

        // Llenar los conjuntos con los datos
        DataFormatter formatter = new DataFormatter();

        while (rowIterator1.hasNext()) {
            Row row = rowIterator1.next();
            Cell cell = row.getCell(3); // Columna de números en Hoja1
            String numeroBuscar = formatter.formatCellValue(cell);
            numerosHoja1.add(numeroBuscar);
        }

        while (rowIterator2.hasNext()) {
            Row row = rowIterator2.next();
            Cell cellNumero = row.getCell(9); // Columna de números en INFORME SOLICITUDES
            Cell cellNombre = row.getCell(10); // Columna de nombres en INFORME SOLICITUDES

            String numeroBuscar = formatter.formatCellValue(cellNumero);
            String nombre = "";

            // Verificar si la celda del nombre no es nula y asignar valor
            if (cellNombre != null) {
                if (cellNombre.getCellType() == Cell.CELL_TYPE_STRING) {
                    nombre = cellNombre.getStringCellValue();
                } else if (cellNombre.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                    nombre = String.valueOf(cellNombre.getNumericCellValue());
                }
            }

            if (!numeroBuscar.isEmpty() && !nombre.isEmpty()) {
                datosINFORME_SOLICITUDES.add(Collections.singletonMap(numeroBuscar, nombre));
            }
        }

        // Iterar sobre los números de la Hoja1 y buscar coincidencias
        rowIterator1 = ws.iterator();
        while (rowIterator1.hasNext()) {
            Row row = rowIterator1.next();
            Cell cell = row.getCell(3); // Columna de números en Hoja1
            String numeroBuscar = formatter.formatCellValue(cell);
            for (Map<String, String> dato : datosINFORME_SOLICITUDES) {
                if (dato.containsKey(numeroBuscar)) {
                    row.createCell(7).setCellValue(dato.get(numeroBuscar)); // Colocar el nombre en la columna 7
                    break;
                }
            }
        }


        System.out.println("Nombres completos agregados exitosamente.");
    }
    /*
    public static void main(String[] args) throws EncryptedDocumentException, IOException {
        String inputFilePath = "O:/aa/result2.xlsx"; // Ruta del archivo .xls original
        String outputFilePath = "O:/aa/result2.xlsx";
        FileInputStream fileInputStream = new FileInputStream(new File(inputFilePath));
        Workbook wb = WorkbookFactory.create(fileInputStream);
        try {
            simulateBUSCARVHoja1(wb);
            //Guardar archivo 
            FileOutputStream fileOutputStream = new FileOutputStream(new File(outputFilePath));
            wb.write(fileOutputStream);
            fileOutputStream.close();
            wb.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }*/
}