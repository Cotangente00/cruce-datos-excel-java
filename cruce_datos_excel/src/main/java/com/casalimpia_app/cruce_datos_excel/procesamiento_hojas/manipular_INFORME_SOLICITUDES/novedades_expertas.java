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
public class novedades_expertas {
    public static void resaltarNovedad(Workbook wb) throws IOException {
        
        Sheet ws = wb.getSheetAt(0);

        // Crear un estilo de celda con relleno amarillo
        CellStyle yellowStyle = wb.createCellStyle();
        yellowStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        yellowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // Iterar sobre las filas (empezando desde la fila 1 para saltar el encabezado)
        for (int rowIndex = 1; rowIndex <= ws.getLastRowNum(); rowIndex++) {
            Row row = ws.getRow(rowIndex);
            if (row != null) {
                // Obtener la celda de la columna N (índice 13)
                Cell cellN = row.getCell(13); // Columna N = índice 13

                if (cellN != null && cellN.getCellType() == Cell.CELL_TYPE_STRING) {
                    String valorNovedad = cellN.getStringCellValue();

                    // Si el valor es "Si", resaltar las celdas en las columnas J (índice 9) y K (índice 10)
                    if (valorNovedad.equalsIgnoreCase("Si")) {
                        Cell cellJ = row.getCell(9); // Columna J = índice 9
                        Cell cellK = row.getCell(10); // Columna K = índice 10

                        if (cellJ != null) {
                            cellJ.setCellStyle(yellowStyle); // Aplicar el estilo amarillo a la columna J
                        }

                        if (cellK != null) {
                            cellK.setCellStyle(yellowStyle); // Aplicar el estilo amarillo a la columna K
                        }
                    }
                }
            }
        }
        int EliminarColumnaN = 13; // Índice de la columna (empezando desde 0)
        for (Row fila : ws) {
            if (fila != null && fila.getCell(EliminarColumnaN) != null) {
                fila.removeCell(fila.getCell(EliminarColumnaN));
            }
        }
        System.out.println("Proceso completado. Las celdas de las columnas J y K han sido resaltadas donde corresponda.");
    }
    /*public static void main(String[] args) {
        try {
            // Ruta del archivo de entrada y salida
            String inputFilePath = "O:/programa/cruce-datos-excel-java/result.xlsx";

            // Llamar a la función para resaltar las filas
            resaltarNovedad(inputFilePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }*/
}
