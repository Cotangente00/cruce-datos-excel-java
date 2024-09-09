package manipular_INFORME_SOLICITUDES;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;


public class novedades_expertas {
    public static void resaltarNovedad(String inputFilePath) throws IOException {
        // Cargar el archivo Excel
        FileInputStream fileInputStream = new FileInputStream(inputFilePath);
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);

        // Crear un estilo de celda con relleno amarillo
        CellStyle yellowStyle = workbook.createCellStyle();
        yellowStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        yellowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // Iterar sobre las filas (empezando desde la fila 1 para saltar el encabezado)
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                // Obtener la celda de la columna N (índice 13)
                Cell cellN = row.getCell(13); // Columna N = índice 13

                if (cellN != null && cellN.getCellType() == CellType.STRING) {
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
        for (Row fila : sheet) {
            if (fila != null && fila.getCell(EliminarColumnaN) != null) {
                fila.removeCell(fila.getCell(EliminarColumnaN));
            }
        }
        

        // Guardar los cambios en un nuevo archivo
        FileOutputStream fileOutputStream = new FileOutputStream(inputFilePath);
        workbook.write(fileOutputStream);

        // Cerrar recursos
        fileOutputStream.close();
        workbook.close();
        fileInputStream.close();

        System.out.println("Proceso completado. Las celdas de las columnas J y K han sido resaltadas donde corresponda.");
    }
    public static void main(String[] args) {
        try {
            // Ruta del archivo de entrada y salida
            String inputFilePath = "O:/programa/cruce-datos-excel-java/result.xlsx";

            // Llamar a la función para resaltar las filas
            resaltarNovedad(inputFilePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}