package manipular_INFORME_SOLICITUDES;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class eliminar_filas {
    public static void main(String[] args) {
        String excelFilePath = "O:/aa/test-lunes-jueves.xlsx"; // Ruta al archivo
        String newExcelFilePath = "O:/aa/result.xlsx"; // Ruta al nuevo archivo

        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook wb = new XSSFWorkbook(fis)) {

            // Selecciona la primera hoja índice cero, o la hoja INFORME SOLICITUDES
            Sheet ws = wb.getSheetAt(0);

            // Itera, elimina las 4 primeras filas incluyendo las filas que están completamente vacías 
            for (int i = 0; i < 4; i++) {
                Row row = ws.getRow(i);
                if (row != null) {
                    ws.removeRow(row);
                }
            }

            // Una vez eliminadas las filas, se recorre las filas para reajustar el índice 
            for (int i = 4; i <= ws.getLastRowNum(); i++) {
                Row row = ws.getRow(i);
                if (row != null) {
                    ws.shiftRows(i, i, -4);
                }
            }

            // Escribe los cambios en el archivo Excel
            try (FileOutputStream fos = new FileOutputStream(newExcelFilePath)) {
                wb.write(fos);
            }

            System.out.println("Las primeras 4 filas han sido eliminadas con éxito.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}