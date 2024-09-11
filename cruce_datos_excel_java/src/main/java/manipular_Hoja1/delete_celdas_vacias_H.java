package manipular_Hoja1;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class delete_celdas_vacias_H {
    public static void limpiar_caracteres_invisibles(String inputFilePath) throws IOException {
        // Cargar el archivo Excel
        FileInputStream fileInputStream = new FileInputStream(inputFilePath);
        try (Workbook wb = new XSSFWorkbook(fileInputStream)) {
            Sheet ws = wb.getSheetAt(1);

            int primeraFilaHoja1 = 4;

            // Recorrer las filas de la hoja "Hoja1" desde la fila 5 en adelante
            for (int i = primeraFilaHoja1; i <= ws.getLastRowNum(); i++) {
                Row filaHoja1 = ws.getRow(i);

                if (filaHoja1 != null) {
                    Cell celdaH = filaHoja1.getCell(7); 

                    if (celdaH.getCellType() == CellType.STRING) {
                        String contenido = celdaH.getStringCellValue().trim().replaceAll("\\s+", "");
                        if (!contenido.isEmpty()) {
                            celdaH.setCellValue(contenido);
                        } else {
                            filaHoja1.removeCell(celdaH);
                        }
                    } 

                }
            }
            FileOutputStream fileOutputStream = new FileOutputStream(inputFilePath);
            wb.write(fileOutputStream);

            // Cerrar recursos
            fileOutputStream.close();
            wb.close();
            fileInputStream.close();

            System.out.println("Caracteres invisibles eliminados exitosamente.");

        }
    }
    public static void main(String[] args) {
        String inputFilePath = "O:/programa/cruce-datos-excel-java/result.xlsx";
        try {
            limpiar_caracteres_invisibles(inputFilePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
