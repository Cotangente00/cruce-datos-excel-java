package manipular_Hoja1;

import org.apache.poi.ss.usermodel.*;
import java.io.IOException;

public class find_table {
    public static void encontrar_tabla(Workbook wb) throws IOException {

        Sheet originalSheet = wb.getSheetAt(1); 

        // Crear una nueva hoja para copiar los datos
        Sheet newSheet = wb.createSheet("Hoja2");

        // Buscar la tabla de 11 columnas en la hoja original
        int startRow = -1;
        int startCol = -1;

        for (Row row : originalSheet) {
            int numColumns = 0;
            for (Cell cell : row) {
                if (!cell.toString().trim().isEmpty()) {
                    numColumns++;
                }
            }
            if (numColumns == 10) { // Se encontró una fila con 11 columnas
                startRow = row.getRowNum();
                startCol = row.getFirstCellNum();
                break;
            }
        }

        if (startRow == -1) {
            System.out.println("No se encontró la tabla de 11 columnas.");
            return;
        }

        // Copiar la tabla desde la hoja original a partir de D5 en la nueva hoja
        int newStartRow = 4; // D5 es la fila 5, en base 0 es fila 4
        int newStartCol = 3; // Columna D es la columna 3 en base 0

        for (int i = startRow; i <= originalSheet.getLastRowNum(); i++) {
            Row originalRow = originalSheet.getRow(i);
            if (originalRow != null) {
                Row newRow = newSheet.createRow(newStartRow++);
                for (int j = 0; j < 11; j++) {
                    Cell originalCell = originalRow.getCell(startCol + j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    Cell newCell = newRow.createCell(newStartCol + j);

                    // Copiar el valor de la celda
                    copyCell(originalCell, newCell);
                }
            }
        }

        // Eliminar la hoja original
        int sheetIndex = wb.getSheetIndex(originalSheet);
        wb.removeSheetAt(sheetIndex);
        // Cambiar el nombre de la hoja a "Hoja1"
        wb.setSheetName(wb.getSheetIndex(newSheet), "Hoja1");

        System.out.println("Tabla transferida correctamente.");
    }
    

    // Método para copiar el contenido de una celda
    private static void copyCell(Cell originalCell, Cell newCell) {
        switch (originalCell.getCellType()) {
            case STRING:
                newCell.setCellValue(originalCell.getStringCellValue());
                break;
            case NUMERIC:
                newCell.setCellValue(originalCell.getNumericCellValue());
                break;
            case BOOLEAN:
                newCell.setCellValue(originalCell.getBooleanCellValue());
                break;
            case FORMULA:
                newCell.setCellFormula(originalCell.getCellFormula());
                break;
            case BLANK:
                newCell.setBlank();
                break;
            case ERROR:
                newCell.setCellErrorValue(originalCell.getErrorCellValue());
                break;
            default:
                break;
        }
    }

    /*public static void main(String[] args) {
        String archivo = "O:/programa/cruce-datos-excel-java/test-lunes-jueves.xlsx";
        try {
            encontrar_tabla(archivo);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }*/
}


