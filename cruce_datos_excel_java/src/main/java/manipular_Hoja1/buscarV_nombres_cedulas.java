package manipular_Hoja1;

import org.apache.poi.ss.usermodel.*;
import java.util.*;
import java.io.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class buscarV_nombres_cedulas {
    public static void simulateBUSCARV(String inputFilePath) throws Exception {
        // Cargar el archivo Excel
        FileInputStream fileInputStream = new FileInputStream(inputFilePath);
        try (Workbook wb = new XSSFWorkbook(fileInputStream)) {
            Sheet ws = wb.getSheetAt(0);
            Sheet ws2 = wb.getSheetAt(1);

            // Obtener las columnas de interés como iteradores
            Iterator<Row> rowIterator1 = ws.iterator();
            rowIterator1.next(); // Saltar el encabezado
            Iterator<Row> rowIterator2 = ws2.iterator();

            // Crear conjuntos para almacenar los números de documento
            Set<String> numerosHoja1 = new HashSet<>();
            Set<Map<String, String>> datosHoja2 = new HashSet<>();

            // Llenar los conjuntos con los datos
            while (rowIterator1.hasNext()) {
                Row row = rowIterator1.next();
                Cell cell = row.getCell(9);
                DataFormatter formatter = new DataFormatter();
                String numeroBuscar = formatter.formatCellValue(cell);
                numerosHoja1.add(numeroBuscar);
            }
            while (rowIterator2.hasNext()) {
                Row row = rowIterator2.next();
                Cell cell = row.getCell(3);
                DataFormatter formatter = new DataFormatter();
                String numeroBuscar = formatter.formatCellValue(cell);
                String nombre = row.getCell(4).getStringCellValue();
                datosHoja2.add(Map.of(numeroBuscar, nombre));
            }

            // Iterar sobre los números de la Hoja1 y buscar coincidencias
            rowIterator1 = ws.iterator();
            rowIterator1.next(); // Saltar el encabezado
            while (rowIterator1.hasNext()) {
                Row row = rowIterator1.next();
                Cell cell = row.getCell(9);
                DataFormatter formatter = new DataFormatter();
                String numeroBuscar = formatter.formatCellValue(cell);
                for (Map<String, String> dato : datosHoja2) {
                    if (dato.containsKey(numeroBuscar)) {
                        row.createCell(12).setCellValue(numeroBuscar);
                        row.createCell(13).setCellValue(dato.get(numeroBuscar));
                        break;
                    }
                }
            }

            int[] columnas = {12};

            for (int rowIndex = 1; rowIndex <= ws.getLastRowNum(); rowIndex++) { // Inicia en 1 para saltar el encabezado
                Row row = ws.getRow(rowIndex);
                if (row != null) {
                    for (int colIndex : columnas) {
                        Cell cell = row.getCell(colIndex);
                        if (cell != null && cell.getCellType() == CellType.STRING) {
                            String cellValue = cell.getStringCellValue();

                            // Verificar si el valor de la celda es numérico o contiene espacios al inicio o final
                            if (cellValue.matches("\\s*\\d+\\s*")) {
                                // Eliminar espacios en blanco y convertir a numérico
                                double numericValue = Double.parseDouble(cellValue.trim());
                                cell.setCellValue(numericValue);
                            }
                        }
                    }
                }
            }

            // Guardar los cambios en un nuevo archivo

            FileOutputStream fileOutputStream = new FileOutputStream(inputFilePath);
            wb.write(fileOutputStream);

            // Cerrar recursos
            fileOutputStream.close();
            wb.close();
            fileInputStream.close();

            System.out.println("Números de documento y nombres completos agregados exitosamente.");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        String rutaArchivo = "O:/programa/cruce-datos-excel-java/result.xlsx";
        try {
            simulateBUSCARV(rutaArchivo);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
