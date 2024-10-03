/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.cruce_datos_excel.procesamiento_hojas.manipular_INFORME_SOLICITUDES;

import org.apache.poi.ss.usermodel.*;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

/**
 *
 * @author jcavilaa
 */
public class filtrar_ciudades {
    public static void filtrarCiudades(Workbook wb) throws IOException {
        // Lista de ciudades válidas
        List<String> ciudadesValidas = Arrays.asList("bogotá", "chía", "cota", "cajicá", "soacha", "", "bogota", "chia", "cajica");

        Sheet ws = wb.getSheetAt(0);

        // Índice de la columna "Ciudad" (M es la columna 12, 0-indexed)
        int columnaCiudadIndex = 12;
        int columnaOIndex = 14;

        // Crear un estilo de celda para las filas cuyas ciudades son "Soacha" y vacías
        CellStyle style = wb.createCellStyle();
        Font font = wb.createFont();
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        font.setUnderline(Font.U_SINGLE);
        style.setFont(font);

        // Iterar sobre las filas y eliminar las que no cumplan con el criterio
        for (int rowIndex = ws.getLastRowNum(); rowIndex >= 1; rowIndex--) {  // Empieza desde el final para evitar problemas con el desplazamiento de filas y saltando el encabezado
            Row row = ws.getRow(rowIndex);
            if (row != null) {
                Cell cellCiudad = row.getCell(columnaCiudadIndex);
                String valorCiudad = (cellCiudad != null) ? cellCiudad.getStringCellValue().trim() : "";
                if (valorCiudad.equalsIgnoreCase("soacha")) {
                    Cell cellColumnaO = row.getCell(columnaOIndex);
                    if (cellColumnaO == null) {
                        cellColumnaO = row.createCell(columnaOIndex);
                    }
                    cellColumnaO.setCellValue("Soacha(Validar Servicio)");
                    cellColumnaO.setCellStyle(style);
                } else if (valorCiudad.isEmpty() || valorCiudad.equalsIgnoreCase("") || valorCiudad == null) {
                    Cell cellColumnaO = row.getCell(columnaOIndex);
                    if (cellColumnaO == null) {
                        cellColumnaO = row.createCell(columnaOIndex);
                    }
                    cellColumnaO.setCellValue("Ciudad vacía(Confirmar)");
                    cellColumnaO.setCellStyle(style);
                } else if (!ciudadesValidas.contains(valorCiudad.toLowerCase())) {
                    int lastRow = ws.getLastRowNum();
                    if (rowIndex < lastRow) {
                        ws.shiftRows(rowIndex + 1, lastRow, -1);
                    } else {
                        ws.removeRow(row);
                    }
                }
            }
        }

        // Eliminar la columna N (índice 12)
        int eliminarColumnaN = 12;  // Índice de la columna (empezando desde 0)
        for (Row fila : ws) {
            if (fila != null && fila.getCell(eliminarColumnaN) != null) {
                fila.removeCell(fila.getCell(eliminarColumnaN));
            }
        }

        System.out.println("Proceso completado. Filas filtradas.");

        // Eliminar todas las filas cuyo valor en la columna A esté vacío
        for (int rowIndex2 = ws.getLastRowNum(); rowIndex2 >= 1; rowIndex2--) {
            Row row = ws.getRow(rowIndex2);
            if (row == null || row.getCell(0) == null || row.getCell(0).getCellType() == Cell.CELL_TYPE_BLANK) {
                ws.removeRow(row);
            }
        }
    }
}
