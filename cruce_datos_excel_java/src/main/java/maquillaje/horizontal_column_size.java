package maquillaje;

import org.apache.poi.ss.usermodel.*;

import java.io.IOException;

public class horizontal_column_size {
    public static void ajustarAnchoColumnas(Workbook wb) throws IOException {

        Sheet ws = wb.getSheetAt(0); // Obteniendo la primera hoja
        Sheet ws2 = wb.getSheetAt(1); // Obteniendo la segunda hoja

        // Ajustar automáticamente todas las columnas con el método autoSizeColumn() INFORME SOLICITUDES
        if (ws.getPhysicalNumberOfRows() > 0) {
            Row primeraFila = ws.getRow(0);

            if (primeraFila != null) {
                int numColumnas = primeraFila.getPhysicalNumberOfCells();

                // Ajustar todas las columnas automáticamente
                for (int colIndex = 0; colIndex < numColumnas; colIndex++) {
                    ws.autoSizeColumn(colIndex);
                }
            }
        }

        int primeraFilaHoja1 = 4;
        // Recorrer las filas de la hoja "Hoja1" desde la fila 5 en adelante
        for (int i = primeraFilaHoja1; i <= ws2.getLastRowNum(); i++) {
            Row filaHoja1 = ws2.getRow(i);

            if (filaHoja1 != null) {
                Cell celdaG = filaHoja1.getCell(6); 
                Cell celdaI = filaHoja1.getCell(8); 
                Cell celdaJ = filaHoja1.getCell(9); 
                Cell celdaK = filaHoja1.getCell(10); 
                Cell celdaL = filaHoja1.getCell(11); 
                Cell celdaM = filaHoja1.getCell(12); 
                Cell celdaN = filaHoja1.getCell(13); 

                if (celdaG.getStringCellValue() == null || celdaG.getStringCellValue().isEmpty()) {
                    filaHoja1.removeCell(celdaG);
                    filaHoja1.removeCell(celdaI);
                    filaHoja1.removeCell(celdaJ);
                    filaHoja1.removeCell(celdaK);
                    filaHoja1.removeCell(celdaL);
                    filaHoja1.removeCell(celdaM);
                    filaHoja1.removeCell(celdaN);
                }
            } 
        }   

        // Ajustar manualmente el ancho de las columnas M (12), N (13), y O (14)
        ajustarColumnasManualmente(ws, 12); // Columna M (índice 12)
        ajustarColumnasManualmente(ws, 13); // Columna N (índice 13)
        ajustarColumnasManualmente(ws, 14); // Columna O (índice 14)
        ajustarColumnasManualmente(ws2, 3); // Columna D (índice 3)
        ajustarColumnasManualmente(ws2, 4); // Columna E (índice 4)
        ajustarColumnasManualmente(ws2, 5); // Columna F (índice 5)
        ajustarColumnasManualmente(ws2, 7); // Columna G (índice 6)     

        System.out.println("Proceso completado. Se ajustó el ancho de las columnas.");
    }

    // Método auxiliar para ajustar manualmente el ancho de columnas con celdas vacías o sin encabezado
    public static void ajustarColumnasManualmente(Sheet sheet, int colIndex) {
        int maxWidth = 0;

        // Multiplicador para ajustar mejor el tamaño en base al contenido
        double widthMultiplier = 1.3; // Mejora el ajuste, considerando mayúsculas

        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) { // Inicia en la fila 1 (segunda fila)
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                Cell cell = row.getCell(colIndex);
                if (cell != null) {
                    String cellValue = cell.toString();
                    int cellWidth = (int) (cellValue.length() * 256 * widthMultiplier); // Convertir el número de caracteres a unidades de ancho de columna

                    // Actualizar el máximo ancho si este valor es mayor
                    if (cellWidth > maxWidth) {
                        maxWidth = cellWidth;
                    }
                }
            }
        }
        // Ajustar el ancho de la columna al valor máximo encontrado
        sheet.setColumnWidth(colIndex, Math.min(maxWidth, 255 * 256)); // Limitar el ancho máximo permitido por Excel
    }
    /*public static void main(String[] args) {
        String inputFilePath = "O:/programa/cruce-datos-excel-java/result.xlsx";
        //String outputFilePath = "O:/programa/cruce-datos-excel-java/result2.xlsx";

        try {
            ajustarAnchoColumnas(inputFilePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }*/
}
