/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.cruce_datos_excel.procesamiento_hojas.maquillaje;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;

import java.io.*;

/**
 *
 * @author jcavilaa
 */
public class delete_image {
    public static void copiarContenidoHoja(Workbook wb) throws IOException {
        // Cargar el archivo Excel
        ZipSecureFile.setMinInflateRatio(0);

        // Obtener la hoja de origen y la hoja de destino
        Sheet ws = wb.getSheetAt(0);
        Sheet wsDestino = wb.createSheet("HojaCopia");

        // Limpiar la hoja de destino antes de copiar los datos
        limpiarHoja(wsDestino);

        // Copiar contenido de la hoja de origen a la hoja de destino, ignorando filas y columnas vacías
        for (Row filaOrigen : ws) {
            if (!filaEstaVacia(filaOrigen)) {
                // Crear o obtener la fila en la hoja de destino
                Row filaDestino = wsDestino.createRow(filaOrigen.getRowNum());

                for (Cell celdaOrigen : filaOrigen) {
                    if (!celdaEstaVacia(celdaOrigen)) {
                        // Crear o obtener la celda en la hoja de destino
                        Cell celdaDestino = filaDestino.createCell(celdaOrigen.getColumnIndex());
                        copiarCelda(celdaOrigen, celdaDestino);
                    }
                }
            }
        }



        // Ajustar automáticamente todas las columnas con el método autoSizeColumn() INFORME SOLICITUDES
        if (wsDestino.getPhysicalNumberOfRows() > 0) {
            Row primeraFila = wsDestino.getRow(0);

            if (primeraFila != null) {
                int numColumnas = primeraFila.getPhysicalNumberOfCells();

                // Ajustar todas las columnas automáticamente
                for (int colIndex = 0; colIndex < numColumnas; colIndex++) {
                    wsDestino.autoSizeColumn(colIndex);
                }       
            }
        }   

        // Ajustar manualmente el ancho de las columnas M (12), N (13), y O (14)
        horizontal_column_size.ajustarColumnasManualmente(wsDestino, 12); // Columna M (índice 12)
        horizontal_column_size.ajustarColumnasManualmente(wsDestino, 13); // Columna N (índice 13)
        horizontal_column_size.ajustarColumnasManualmente(wsDestino, 14); // Columna O (índice 14)



        //Eliminar la hoja original
        wb.removeSheetAt(0);



        // Reordenar las hojas del archivo
        Sheet ws1 = (Sheet) wb.getSheetAt(0);
        int indiceHoja1 = wb.getSheetIndex(ws1);
        // Mover la hoja "INFORME SOLICITUDES" al índice 1 (después de "Hoja1")
        wb.setSheetOrder(ws1.getSheetName(), indiceHoja1 + 1);
        // Renombrar la hoja
        wb.setSheetName(0, "INFORME SOLICITUDES");

    }

    // Función para copiar el contenido de una celda
    private static void copiarCelda(Cell celdaOrigen, Cell celdaDestino) {
        // Copiar el valor según el tipo de la celda
        switch (celdaOrigen.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                celdaDestino.setCellValue(celdaOrigen.getStringCellValue());
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(celdaOrigen)) {
                    celdaDestino.setCellValue(celdaOrigen.getDateCellValue());
                } else {
                    celdaDestino.setCellValue(celdaOrigen.getNumericCellValue());
                }
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                celdaDestino.setCellValue(celdaOrigen.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_FORMULA:
                celdaDestino.setCellFormula(celdaOrigen.getCellFormula());
                break;
            case Cell.CELL_TYPE_BLANK:
                celdaDestino.setCellType(null);
                break;
            case Cell.CELL_TYPE_ERROR:
                celdaDestino.setCellErrorValue(celdaOrigen.getErrorCellValue());
            default:
                break;
        }

        // Copiar el estilo de la celda
        CellStyle estiloOrigen = celdaOrigen.getCellStyle();
        CellStyle estiloDestino = celdaDestino.getSheet().getWorkbook().createCellStyle();
        estiloDestino.cloneStyleFrom(estiloOrigen);
        celdaDestino.setCellStyle(estiloDestino);
    }

    // Función para verificar si una fila está vacía
    private static boolean filaEstaVacia(Row fila) {
        for (Cell celda : fila) {
            if (!celdaEstaVacia(celda)) {
                return false;
            }
        }
        return true;
    }

    // Función para verificar si una celda está vacía
    private static boolean celdaEstaVacia(Cell celda) {
        if (celda == null || celda.getCellType() == Cell.CELL_TYPE_BLANK) {
            return true;
        }
        if (celda.getCellType() == Cell.CELL_TYPE_STRING && celda.getStringCellValue().trim().isEmpty()) {
            return true;
        }
        return false;
    }

    // Función para limpiar la hoja de destino antes de copiar los datos
    private static void limpiarHoja(Sheet hojaDestino) {
        for (int i = hojaDestino.getLastRowNum(); i >= 0; i--) {
            Row fila = hojaDestino.getRow(i);
            if (fila != null) {
                hojaDestino.removeRow(fila);
            }
        }
    }

    public static void main(String[] args) throws Exception {
        String inputFilePath = "O:/aa/result2.xlsx";
        String outputFilePath = "O:/aa/result3.xlsx";
        ZipSecureFile.setMinInflateRatio(0);
        Workbook wb;
        try (FileInputStream fis = new FileInputStream(new File(inputFilePath))) {
            wb = WorkbookFactory.create(fis);  // Apache POI detecta automáticamente si es .xls o .xlsx
        }

        try {
            copiarContenidoHoja(wb);
            wb.write(new FileOutputStream(outputFilePath));
            wb.close();
            System.out.println("Archivo procesado exitosamente.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
