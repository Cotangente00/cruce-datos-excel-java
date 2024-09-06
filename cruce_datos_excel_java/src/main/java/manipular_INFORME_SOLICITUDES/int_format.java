package manipular_INFORME_SOLICITUDES;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class int_format {
    public static void convertirTextoANumero(String inputFilePath) throws IOException {
        // Cargar el archivo Excel
        FileInputStream fileInputStream = new FileInputStream(inputFilePath);
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);

        // Convertir columnas A (índice 0), B (índice 1) y J (índice 9)
        int[] columnas = {0, 1, 9};

        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) { // Inicia en 1 para saltar el encabezado
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                for (int colIndex : columnas) {
                    Cell cell = row.getCell(colIndex);
                    if (cell != null && cell.getCellType() == CellType.STRING) {
                        String cellValue = cell.getStringCellValue();

                        // Verificar si el valor de la celda es numérico
                        if (cellValue.matches("\\d+")) {
                            // Convertir el valor de la celda a numérico
                            double numericValue = Double.parseDouble(cellValue);

                            // Cambiar el tipo de celda a numérico
                            cell.setCellValue(numericValue);
                        }
                    }
                }
            }
        }

        // Guardar los cambios en un nuevo archivo
        FileOutputStream fileOutputStream = new FileOutputStream(inputFilePath);
        workbook.write(fileOutputStream);

        // Cerrar recursos
        fileOutputStream.close();
        workbook.close();
        fileInputStream.close();

        System.out.println("Proceso completado. Datos convertidos a formato numérico y archivo guardado en: " + inputFilePath);
    }
}