package procesamiento_hojas.manipular_Hoja1_weekend;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.*;

public class concatenar_nombres_apellidos_wknd {
    public static void concatenacion_wknd(Workbook wb) throws IOException {

        Sheet ws = wb.getSheetAt(1); // Obtener la hoja

        int columnaE = 4, columnaF = 5;
        for (Row row : ws) {
            Cell celdaE = row.getCell(columnaE);
            Cell celdaF = row.getCell(columnaF);
            if (celdaE != null && celdaF != null) {
                String nombre = celdaE.getStringCellValue();
                String apellido = celdaF.getStringCellValue();
                String nombreCompleto = nombre + " " + apellido;
                celdaE.setCellValue(nombreCompleto);
            }   
        }
        int[] columnasEliminar = {5,6,7,8,9,10,11,12,13,15};
        // Recorrer todas las filas de la hoja
        for (Row fila : ws) {
            // Recorrer las columnas a eliminar en orden inverso para evitar problemas
            for (int i = columnasEliminar.length - 1; i >= 0; i--) {
                int columna = columnasEliminar[i];

                // Mover todas las celdas a la izquierda
                for (int j = columna; j < fila.getLastCellNum() - 1; j++) {
                    Cell celdaActual = fila.getCell(j);
                    Cell celdaSiguiente = fila.getCell(j + 1);

                    if (celdaSiguiente != null) {
                        if (celdaActual == null) {
                            celdaActual = fila.createCell(j);
                        }
                        copiarCelda(celdaSiguiente, celdaActual);
                    } else if (celdaActual != null) {
                        fila.removeCell(celdaActual);
                    }
                }
            }
        }
        
        int[] columasABorrar = {7,8,9,10,11,12,13,14,15};
        for (Row row : ws) {
            // Iteramos sobre las columnas a borrar
            for (int col : columasABorrar) {
                Cell cell = row.getCell(col);
                if (cell != null) {
                    cell.setCellValue(""); // Establecemos el valor de la celda como cadena vacía
                }
            }
        }
             
        System.out.println("Nombres y apellidos concatenados exitosamente.");
    }

    // Función para copiar el contenido de una celda a otra sin usar setCellType
    private static void copiarCelda(Cell desde, Cell hacia) {
        switch (desde.getCellType()) {
            case STRING:
                hacia.setCellValue(desde.getStringCellValue());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(desde)) {
                    hacia.setCellValue(desde.getDateCellValue());
                } else {
                    hacia.setCellValue(desde.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                hacia.setCellValue(desde.getBooleanCellValue());
                break;
            case FORMULA:
                hacia.setCellFormula(desde.getCellFormula());
                break;
            case BLANK:
                hacia.setBlank();
                break;
            case ERROR:
                hacia.setCellErrorValue(desde.getErrorCellValue());
                break;
            default:
                   break;
        }
    }

    public static void main(String[] args) throws Exception {
        String inputFilePath = "O:/aa/result.xlsx";
        String outputFilePath = "O:/aa/result2.xlsx";
        Workbook wb;
        try (FileInputStream fis = new FileInputStream(new File(inputFilePath))) {
            wb = WorkbookFactory.create(fis);  // Apache POI detecta automáticamente si es .xls o .xlsx
        }
        try {
            concatenacion_wknd(wb);
            wb.write(new FileOutputStream(outputFilePath));
            wb.close();
            System.out.println("Archivo procesado exitosamente.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
