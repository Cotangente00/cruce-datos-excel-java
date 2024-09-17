package manipular_INFORME_SOLICITUDES;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class eliminar_filas {
    public static void delete_filas(String inputFilePath, String outputFilePath) throws IOException{
        // Cargar archivo Excel
        FileInputStream fileInputStream = new FileInputStream(inputFilePath);
        Workbook wb = new XSSFWorkbook(fileInputStream);
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

        // Escribir los cambios en un archivo nuevo o sobreescribir el original
        FileOutputStream fileOutputStream = new FileOutputStream(outputFilePath);
        wb.write(fileOutputStream);

        // Cerrar recursos
        fileOutputStream.close();
        wb.close();
        fileInputStream.close();

        System.out.println("Las primeras 4 filas han sido eliminadas con éxito.");   
    }
    
    public static void main(String[] args) {
        String inputFilePath = "O:/aa/test-lunes-jueves.xlsx";
        String outputFilePath = "O:/aa/result.xlsx";

        try {
            delete_filas(inputFilePath, outputFilePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}