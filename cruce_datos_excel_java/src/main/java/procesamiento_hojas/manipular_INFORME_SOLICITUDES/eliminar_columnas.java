package procesamiento_hojas.manipular_INFORME_SOLICITUDES;

import org.apache.poi.ss.usermodel.*;
import java.io.IOException;

public class eliminar_columnas {
    public static void eliminarColumnas(Workbook wb) throws IOException {

        Sheet ws = wb.getSheetAt(0); // Obtener la primera hoja

        // Índices de las columnas a eliminar (empiezan desde 0: A=0, B=1, C=2, etc.)
        int[] columnasAEliminar = {2, 3, 9, 10};

        // Recorrer todas las filas de la hoja
        for (Row fila : ws) {
            // Recorrer las columnas a eliminar en orden inverso para evitar problemas
            for (int i = columnasAEliminar.length - 1; i >= 0; i--) {
                int columna = columnasAEliminar[i];

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

    /*public static void main(String[] args) {
        String rutaArchivo = "O:/programa/cruce-datos-excel-java/test-lunes-jueves.xlsx";

        try {
            eliminarColumnas(rutaArchivo);
            System.out.println("Columnas eliminadas correctamente.");
        } catch (IOException e) {
            System.out.println("Ocurrió un error al procesar el archivo: " + e.getMessage());
        }
    }*/
}