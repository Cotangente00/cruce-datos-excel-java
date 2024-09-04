package controller;

import org.apache.poi.ss.usermodel.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class manipular_INFORME_SOLICITUDES {

    public static void eliminarPrimerasFilas(String rutaArchivoExcel, String rutaArchivoExcelNuevo) throws IOException {
        // Se crea un objeto para leer el archivo Excel
        FileInputStream file = new FileInputStream(new File(rutaArchivoExcel));
        Workbook workbook = WorkbookFactory.create(file);

        // Se Obtiene la hoja "INFORME SOLICITUDES" con índice cero dentro del código.
        Sheet sheet = workbook.getSheetAt(0);

        // Obtener el número total de filas en la hoja
        int lastRowNum = sheet.getLastRowNum();

        // Iterar solo si hay suficientes filas
        for (int i = 3; i <= lastRowNum && i >= 0; i--) {
            Row row = sheet.getRow(i);
            if (row != null) {
                sheet.removeRow(row);
            }
        }

        // Se crea un objeto para escribir los cambios en el archivo
        FileOutputStream outputStream = new FileOutputStream(rutaArchivoExcelNuevo);
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
    }

    public static void main(String[] args) {
        String rutaArchivo = "O:/programa/cruce-datos-excel-java/test-lunes-jueves-xlsx.xlsx"; // Ruta del archivo entrante
        String rutaArchivoNuevo = "O:/programa/cruce-datos-excel-java/result.xlsx"; // Ruta del archivo saliente 
        try {
            manipular_INFORME_SOLICITUDES.eliminarPrimerasFilas(rutaArchivo, rutaArchivoNuevo);
            System.out.println("Las primeras 4 filas se han eliminado correctamente.");
        } catch (IOException e) {
            System.out.println("Error al eliminar las filas: " + e.getMessage());
        }
    }
}