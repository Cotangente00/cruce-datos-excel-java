package manipular_INFORME_SOLICITUDES;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

public class filtrar_ciudades {
    public static void filtrarCiudades(String inputFilePath) throws IOException {
        // Lista de ciudades válidas
        List<String> ciudadesValidas = Arrays.asList("bogotá", "chía", "cota", "cajicá", "soacha", "", "bogota", "chia", "cajica");

        // Cargar archivo Excel
        FileInputStream fileInputStream = new FileInputStream(inputFilePath);
        Workbook wb = new XSSFWorkbook(fileInputStream);
        Sheet ws = wb.getSheetAt(0);

        // Encontrar el índice de la columna "Ciudad" (M es la columna 12, 0-indexed)
        int columnaCiudadIndex = 12;
        int columnaOIndex = 14;

        // Iterar sobre las filas y eliminar las que no cumplan con el criterio
        for (int rowIndex = ws.getLastRowNum(); rowIndex >= 1; rowIndex--) {  // Empieza desde el final para evitar problemas con el shift de filas y salteandose el encabezado
            Row row = ws.getRow(rowIndex);
            if (row != null) {
                Cell cellCiudad = row.getCell(columnaCiudadIndex);
                String valorCiudad = (cellCiudad != null) ? cellCiudad.getStringCellValue().trim() : "";
                if (valorCiudad.equalsIgnoreCase("soacha")){
                    Cell cellColumnaO = row.getCell(columnaOIndex);
                    if (cellColumnaO == null) {
                        cellColumnaO = row.createCell(columnaOIndex);
                    }  
                    cellColumnaO.setCellValue("Soacha(Validar Servicio)");
                } else if (valorCiudad.isEmpty()){
                    Cell cellColumnaO = row.getCell(columnaOIndex);
                    if (cellColumnaO == null) {
                        cellColumnaO = row.createCell(columnaOIndex);
                    }  
                    cellColumnaO.setCellValue("Ciudad vacía(Confirmar)");
                } else if (!ciudadesValidas.contains(valorCiudad.toLowerCase())) {
                    int lastRow = ws.getLastRowNum();
                    if (rowIndex < lastRow) {
                        ws.shiftRows(rowIndex + 1, lastRow, -1);
                    } else {
                        ws.removeRow(row);
                    }
                } 
            }
        }

        int EliminarColumnaN = 12; // Índice de la columna (empezando desde 0)
        for (Row fila : ws) {
            if (fila != null && fila.getCell(EliminarColumnaN) != null) {
                fila.removeCell(fila.getCell(EliminarColumnaN));
            }
        }

        // Escribir los cambios en un nuevo archivo
        FileOutputStream fileOutputStream = new FileOutputStream(inputFilePath);
        wb.write(fileOutputStream);

        // Cerrar recursos
        fileOutputStream.close();
        wb.close();
        fileInputStream.close();

        System.out.println("Proceso completado. Filas filtradas y archivo guardado en: " + inputFilePath);
    }
}
