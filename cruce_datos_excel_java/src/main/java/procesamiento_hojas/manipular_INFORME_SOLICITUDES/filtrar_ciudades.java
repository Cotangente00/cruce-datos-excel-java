package procesamiento_hojas.manipular_INFORME_SOLICITUDES;

import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.util.Arrays;
import java.util.List;

public class filtrar_ciudades {
    public static void filtrarCiudades(Workbook wb) throws IOException {
        // Lista de ciudades válidas
        List<String> ciudadesValidas = Arrays.asList("bogotá", "chía", "cota", "cajicá", "soacha", "", "bogota", "chia", "cajica");

        Sheet ws = wb.getSheetAt(0);

        // Encontrar el índice de la columna "Ciudad" (M es la columna 12, 0-indexed)
        int columnaCiudadIndex = 12;
        int columnaOIndex = 14;

        //Crear un estilo de celda para las expertas cuyas ciudades son "Soacha" y NULL
        CellStyle style = wb.createCellStyle();
        Font font = wb.createFont();
        font.setBold(true);
        font.setUnderline(Font.U_SINGLE);
        style.setFont(font);



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
                    cellColumnaO.setCellStyle(style);
                } else if (valorCiudad.isEmpty() || valorCiudad.equalsIgnoreCase("") || valorCiudad == null){
                    Cell cellColumnaO = row.getCell(columnaOIndex);
                    if (cellColumnaO == null) {
                        cellColumnaO = row.createCell(columnaOIndex);
                    }  
                    cellColumnaO.setCellValue("Ciudad vacía(Confirmar)");
                    cellColumnaO.setCellStyle(style);
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
        

        System.out.println("Proceso completado. Filas filtradas.");
        
        //Eliminar todas las filas cuyo valor en la columna A esté vacío
        for (int rowIndex2 = ws.getLastRowNum(); rowIndex2 >= 1; rowIndex2--) {
            Row row = ws.getRow(rowIndex2);
            if (row == null || row.getCell(0) == null || row.getCell(0).getCellType() == CellType.BLANK) {
                ws.removeRow(row);
            }
        }
    }
}
