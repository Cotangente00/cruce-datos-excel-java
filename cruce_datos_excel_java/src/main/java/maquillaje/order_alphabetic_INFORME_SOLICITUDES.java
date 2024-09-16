package maquillaje;

    import org.apache.poi.ss.usermodel.*;
    import org.apache.poi.xssf.usermodel.XSSFWorkbook;

    import java.io.FileInputStream;
    import java.io.FileOutputStream;
    import java.io.IOException;
    import java.util.ArrayList;
    import java.util.List;

public class order_alphabetic_INFORME_SOLICITUDES {
    public static void reorganizeExcel_INFORME_SOLICITUDES(String inputFilePath) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(inputFilePath);
        Workbook wb = new XSSFWorkbook(fileInputStream);
        Sheet originalSheet = wb.getSheetAt(0);  // Obtener la primera hoja
        Sheet newSheet = wb.createSheet("ReorganizedSheet");  // Crear una nueva hoja para los datos reorganizados

        List<Row> rowsWithEmptyMN = new ArrayList<>();
        List<Row> rowsWithData = new ArrayList<>();
        List<Row> data = new ArrayList<>();

        // Iterar sobre la columna A para determinar el rango de filas
        int rowIndex = 1;  // Empezar desde la fila 2 (índice 1)
        while (true) {
            Row row = originalSheet.getRow(rowIndex);
            if (row == null || row.getCell(0) == null || row.getCell(0).getCellType() == CellType.BLANK) {
                break;  // Detener cuando se encuentre la primera celda vacía en la columna A
            }

            // Verificar las celdas en las columnas M y N (índices 12 y 13 respectivamente)
            Cell cellM = row.getCell(12);
            Cell cellN = row.getCell(13);

            boolean isCellMEmpty = (cellM == null || cellM.getCellType() == CellType.BLANK);
            boolean isCellNEmpty = (cellN == null || cellN.getCellType() == CellType.BLANK);

            // Si ambas celdas (M y N) están vacías, agregar la fila a la lista correspondiente
            if (isCellMEmpty && isCellNEmpty) {
                rowsWithEmptyMN.add(row);
            } else {
                rowsWithData.add(row);
            }

            rowIndex++;
        }

        // Copiar filas con datos primero
        int newRowIndex = 1;  // Comenzar desde la fila 2 en la nueva hoja
        for (Row row : rowsWithData) {
            copyRow(row, newSheet.createRow(newRowIndex++), wb);
        }

        // Copiar filas con celdas vacías en M y N al final
        for (Row row : rowsWithEmptyMN) {
            copyRow(row, newSheet.createRow(newRowIndex++), wb);
        }

        //eliminar datos de la hoja original omitiendo los encabezados
        for (int rowIndex2 = 1; rowIndex2 <= originalSheet.getLastRowNum(); rowIndex2++) {
            Row row = originalSheet.getRow(rowIndex2);
            if (row != null) {
                originalSheet.removeRow(row);
            } else {
                break;
            }
        }


        // Almacenar todos los datos de la columna
        int rowIndex2 = 1;  // Comenzar desde la fila 2 en la hoja nueva
        while (true) {
            Row row = newSheet.getRow(rowIndex2);
            if (row == null || row.getCell(0) == null || row.getCell(0).getCellType() == CellType.BLANK) {
                break;  // Detener cuando se encuentre la primera celda vacía en la columna A
            }
            data.add(row);
            rowIndex2++;
        }   

        int newRowIndex2 = 1;  // Comenzar desde la fila 2 en la nueva hoja
        for (Row row : data) {
            copyRow(row, originalSheet.createRow(newRowIndex2++), wb);
        }

        wb.removeSheetAt(2);


        // Guardar el archivo de salida
        FileOutputStream fileOutputStream = new FileOutputStream(inputFilePath);
        wb.write(fileOutputStream);
        fileOutputStream.close();
        wb.close();
    }

    // Método para copiar el contenido de una fila a otra
    public static void copyRow(Row sourceRow, Row targetRow, Workbook wb) {
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            Cell sourceCell = sourceRow.getCell(i);
            Cell targetCell = targetRow.createCell(i);

            if (sourceCell != null) {
                switch (sourceCell.getCellType()) {
                    case STRING:
                        targetCell.setCellValue(sourceCell.getStringCellValue());
                        break;
                    case NUMERIC:
                        targetCell.setCellValue(sourceCell.getNumericCellValue());
                        break;
                    case BOOLEAN:
                        targetCell.setCellValue(sourceCell.getBooleanCellValue());
                        break;
                    case FORMULA:
                        targetCell.setCellFormula(sourceCell.getCellFormula());
                        break;
                    default:
                        break;
                }

                CellStyle newCellStyle = wb.createCellStyle();
                newCellStyle.cloneStyleFrom(sourceCell.getCellStyle());
                targetCell.setCellStyle(newCellStyle);
            }
        }
    }




    public static void main(String[] args) {
        try {
            String inputFilePath = "O:/aa/result.xlsx";
            //String outputFilePath = "O:/aa/result2.xlsx";
            reorganizeExcel_INFORME_SOLICITUDES(inputFilePath);
            System.out.println("Archivo reorganizado exitosamente.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
