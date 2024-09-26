package procesamiento_hojas.manipular_INFORME_SOLICITUDES;


import java.io.IOException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;


public class eliminar_filas {
    public static void delete_filas(Workbook wb) throws IOException{
        // Cargar la hoja
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

        System.out.println("Las primeras 4 filas han sido eliminadas con éxito.");   
    }
    
    /*public static void main(String[] args) {
        String inputFilePath = "O:/aa/test-lunes-jueves.xlsx";
        String outputFilePath = "O:/aa/result.xlsx";

        try {
            delete_filas(inputFilePath, outputFilePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }*/
}