/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.cruce_datos_excel.procesamiento_hojas.manipular_INFORME_SOLICITUDES;

import org.apache.poi.ss.usermodel.*;

import java.io.IOException;

/**
 *
 * @author jcavilaa
 */
public class int_format {
    public static void convertirTextoANumero(Workbook wb) throws IOException {

        Sheet ws = wb.getSheetAt(0);

        // Convertir columnas A (índice 0), B (índice 1) y J (índice 9)
        int[] columnas = {0, 1, 9};

        for (int rowIndex = 1; rowIndex <= ws.getLastRowNum(); rowIndex++) { // Inicia en 1 para saltar el encabezado
            Row row = ws.getRow(rowIndex);
            if (row != null) {
                for (int colIndex : columnas) {
                    Cell cell = row.getCell(colIndex);
                    if (cell != null && cell.getCellType() == Cell.CELL_TYPE_STRING) {
                        String cellValue = cell.getStringCellValue();

                        // Verificar si el valor de la celda es numérico o contiene espacios al inicio o final
                        if (cellValue.matches("\\s*\\d+\\s*")) {
                            // Eliminar espacios en blanco y convertir a numérico
                            double numericValue = Double.parseDouble(cellValue.trim());
                            cell.setCellValue(numericValue);
                        }
                    }
                }
            }
        }
        System.out.println("Proceso completado. Datos convertidos a formato numérico.");
    }
}
