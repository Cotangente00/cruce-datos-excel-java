package manipular_Hoja1;

import org.apache.poi.ss.usermodel.*;

import java.text.DecimalFormat;

public class no_service_copy_paste {
    public static void copiarFilas(Workbook wb) throws Exception {
        Sheet wsINFORME_SOLICITUDES = wb.getSheetAt(0);
        Sheet wsHoja1 = wb.getSheetAt(1);

        if (wsHoja1 == null || wsINFORME_SOLICITUDES == null) {
            System.out.println("Una de las hojas no existe.");
            return;
        } 

        // Formato para preservar números sin notación científica
        DecimalFormat df = new DecimalFormat("0");

        // Obtener la última fila de la tabla en Hoja1 (A1:L*)
        int ultimaFilawsINFORME_SOLICITUDES = obtenerUltimaFilaTabla(wsINFORME_SOLICITUDES);

        // La primera fila donde comenzaremos a copiar en la columna M es 10 filas después de la última fila de la tabla
        int filaDestino = ultimaFilawsINFORME_SOLICITUDES + 10;

        int primeraFilaHoja1 = 4;
        
        // Recorrer las filas de la hoja "Hoja1" desde la fila 5 en adelante
        for (int i = primeraFilaHoja1; i <= wsHoja1.getLastRowNum(); i++) {
            Row filaHoja1 = wsHoja1.getRow(i);

            if (filaHoja1 != null) {
                Cell celdaH = filaHoja1.getCell(7);  // Columna H es el índice 7

                // Si la celda H está vacía, copiar los datos de D, E y F a P, Q y R en Hoja1
                if (celdaH == null || esCeldaVaciaOInvisble(celdaH)) {
                    Row filaINFORME_SOLICITUDES = wsINFORME_SOLICITUDES.getRow(filaDestino);
                    if (filaINFORME_SOLICITUDES == null) {
                        filaINFORME_SOLICITUDES = wsINFORME_SOLICITUDES.createRow(filaDestino);
                    }

                    // Copiar D, E y F a P, Q y R
                    Cell celdaD = filaHoja1.getCell(3);  // Columna D es el índice 3
                    Cell celdaE = filaHoja1.getCell(4);  // Columna E es el índice 4
                    Cell celdaF = filaHoja1.getCell(5);  // Columna F es el índice 5

                    // Crear y asignar valores a las celdas P, Q y R en "INFORME SOLICITUDES"
                    Cell celdaP = filaINFORME_SOLICITUDES.createCell(12);  // Columna M es el índice 12
                    Cell celdaQ = filaINFORME_SOLICITUDES.createCell(13);  // Columna N es el índice 13
                    Cell celdaR = filaINFORME_SOLICITUDES.createCell(14);  // Columna O es el índice 14

                    System.out.println("Copiando fila: " + (i + 1) + " | D:" + (celdaD != null ? celdaD.toString() : "") + " | E:" + (celdaE != null ? celdaE.toString() : "") + " | F:" + (celdaF != null ? celdaF.toString() : ""));

                    // Asignar los valores a P, Q y R si las celdas de D, E, F no están vacías
                    if (celdaD != null) {
                        // Si es numérico, usar formato decimal para evitar notación científica
                        if (celdaD.getCellType() == CellType.NUMERIC) {
                            celdaP.setCellValue(df.format(celdaD.getNumericCellValue()));
                        } else {
                            celdaP.setCellValue(celdaD.toString());
                        }
                    }
                    if (celdaE != null) celdaQ.setCellValue(celdaE.toString());
                    if (celdaF != null) celdaR.setCellValue(celdaF.toString());

                    filaDestino++;  // Mover a la siguiente fila en Hoja1
                }
            }
        }

        // Convertir columnas A (índice 0), B (índice 1) y J (índice 9)
        int[] columna = {12};

        for (int rowIndex = 1; rowIndex <= wsINFORME_SOLICITUDES.getLastRowNum(); rowIndex++) { // Inicia en 1 para saltar el encabezado
            Row row = wsINFORME_SOLICITUDES.getRow(rowIndex);
            if (row != null) {
                for (int colIndex : columna) {
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
        System.out.println("Filas copiadas exitosamente.");
    }

    // Método para verificar si una celda está vacía o contiene caracteres invisibles
    public static boolean esCeldaVaciaOInvisble(Cell celda) {
        if (celda.getCellType() == CellType.STRING) {
            // Limpiar caracteres invisibles
            String contenido = celda.getStringCellValue().trim().replaceAll("\\s+", "");
            return contenido.isEmpty();  // True si está vacía después de eliminar espacios y caracteres no imprimibles
        }
        return false;  // La celda no es de tipo STRING o no está vacía
    }

    // Método para obtener la última fila de la tabla en la hoja "INFORME SOLICITUDES"
    public static int obtenerUltimaFilaTabla(Sheet wsINFORME_SOLICITUDES) {
        int ultimaFila = 0;

        for (int i = 0; i <= wsINFORME_SOLICITUDES.getLastRowNum(); i++) {
            Row fila = wsINFORME_SOLICITUDES.getRow(i);
            if (fila != null) {
                for (int j = 0; j <= 11; j++) {  // Revisar las columnas de A a L (índices 0 a 11)
                    Cell celda = fila.getCell(j);
                    if (celda != null && celda.getCellType() != CellType.BLANK) {
                        ultimaFila = i;  // Actualizar la última fila no vacía
                        break;
                    }
                }
            }
        }

        return ultimaFila;
    }

    /*public static void main(String[] args) {
        String rutaArchivo = "O:/programa/cruce-datos-excel-java/result.xlsx";
        //String rutaArchivoSalida = "O:/programa/cruce-datos-excel-java/result2.xlsx";
        try {
            copiarFilas(rutaArchivo);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }*/
}
