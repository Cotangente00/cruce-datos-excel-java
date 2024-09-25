package funciones_boton;

import java.io.*;
import manipular_INFORME_SOLICITUDES.*;
import manipular_Hoja1.*;
import maquillaje.*;


import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class funciones_lunes_jueves {
    public static void lunes_jueves(String inputFilePath, String outputFilePath) throws Exception {
        try {
            //String inputFilePath = "O:/aa/test-lunes-jueves.xls"; // Ruta del archivo .xls original
            //String outputFilePath = "O:/aa/result.xlsx"; // Ruta de salida del archivo .xlsx
            ZipSecureFile.setMinInflateRatio(0);

            // Detectar el formato del archivo y cargar el libro de trabajo
            Workbook wb;
            try (FileInputStream fis = new FileInputStream(new File(inputFilePath))) {
                wb = WorkbookFactory.create(fis);  // Apache POI detecta automáticamente si es .xls o .xlsx
            }

            // Llamada a las funciones de procesamiento
            eliminar_filas.delete_filas(wb);
            eliminar_columnas.eliminarColumnas(wb);
            borrar_columnas_restantes.borrarColumnasRestantes(wb);
            filtrar_ciudades.filtrarCiudades(wb);
            date_format.formatearFechas(wb);
            int_format.convertirTextoANumero(wb);
            novedades_expertas.resaltarNovedad(wb);
            find_table.encontrar_tabla(wb);
            concatenar_nombres_apellidos.concatenacion(wb);
            buscarV_nombres_cedulas.simulateBUSCARV(wb);
            buscarV_nombres.simulateBUSCARVHoja1(wb);
            delete_celdas_vacias_H.limpiar_caracteres_invisibles(wb);
            // Guardar el archivo en formato XLSX
            if (inputFilePath.endsWith(".xls")) {
                try (FileOutputStream outputStream = new FileOutputStream(outputFilePath)) {
                    Workbook outputWorkbook = new XSSFWorkbook();  // Crear un nuevo libro de trabajo para XLSX

                    // Copiar las hojas del workbook original al nuevo workbook de formato XLSX
                    for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                        outputWorkbook.createSheet(wb.getSheetName(i));
                        copiarHoja(wb.getSheetAt(i), outputWorkbook.getSheetAt(i));
                    }

                    outputWorkbook.write(outputStream);  // Guardar el nuevo archivo en formato XLSX
                    outputWorkbook.close();  // Cerrar el workbook de salida
                }

                wb.close();  // Cerrar el workbook original
                String inputFilePath2 = outputFilePath; // Abrir el archivo .xlsx nuevo
                Workbook wb2;
                try (FileInputStream fis = new FileInputStream(new File(inputFilePath2))) {
                    wb2 = WorkbookFactory.create(fis);  // Apache POI detecta automáticamente si es .xls o .xlsx
                }
                estilos_encabezados_xls.estilos_encabezados(wb2);
                date_format.formatearFechas(wb2);
                no_service_copy_paste.copiarFilas(wb2);
                horizontal_column_size.ajustarAnchoColumnas(wb2);
                order_INFORME_SOLICITUDES.reorganizeExcel_INFORME_SOLICITUDES(wb2);
                order_alphabetic_Hoja1.reorganizeExcel_Hoja1(wb2);
                delete_image.copiarContenidoHoja(wb2);
                //Ajustar la altura de la primera fila
                Sheet ws = wb2.getSheetAt(0); // Obteniendo la primera hoja
                Row fila = ws.getRow(0); //Obteniendo la primera fila de la hoja
                fila.setHeightInPoints(20); //Altura de la fila
                System.out.println("Archivo procesado exitosamente.");
                //Guardar el archivo de manera convencional
                wb2.write(new FileOutputStream(outputFilePath));
                wb2.close();
                //SIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIUUUUUUUUUUUUUUUUUUUUUU working
            } else {
                no_service_copy_paste.copiarFilas(wb);
                horizontal_column_size.ajustarAnchoColumnas(wb);
                order_INFORME_SOLICITUDES.reorganizeExcel_INFORME_SOLICITUDES(wb);
                order_alphabetic_Hoja1.reorganizeExcel_Hoja1(wb);
                delete_image.copiarContenidoHoja(wb);
                System.out.println("Archivo procesado exitosamente.");

                //Guardar el archivo de manera convencional
                wb.write(new FileOutputStream(outputFilePath));
                wb.close();
            }
        } catch (IOException e) {
            System.out.println("Ocurrió un error al procesar el archivo: " + e.getMessage());
        }
    }

    // Función para copiar una hoja de un workbook a otro
    private static void copiarHoja(Sheet hojaOrigen, Sheet hojaDestino) {
        Workbook wbDestino = hojaDestino.getWorkbook();
        Workbook wbOrigen = hojaOrigen.getWorkbook();

        for (int i = 0; i <= hojaOrigen.getLastRowNum(); i++) {
            Row filaOrigen = hojaOrigen.getRow(i);
            Row filaDestino = hojaDestino.createRow(i);

            if (filaOrigen != null) {
                for (int j = 0; j < filaOrigen.getLastCellNum(); j++) {
                    Cell celdaOrigen = filaOrigen.getCell(j);
                    Cell celdaDestino = filaDestino.createCell(j);

                    if (celdaOrigen != null) {
                        switch (celdaOrigen.getCellType()) {
                            case STRING:
                                celdaDestino.setCellValue(celdaOrigen.getStringCellValue());
                                break;
                            case NUMERIC:
                                celdaDestino.setCellValue(celdaOrigen.getNumericCellValue());
                                break;
                            case BOOLEAN:
                                celdaDestino.setCellValue(celdaOrigen.getBooleanCellValue());
                                break;
                            case FORMULA:
                                celdaDestino.setCellFormula(celdaOrigen.getCellFormula());
                                break;
                            default:
                                break;
                        } 
                        CellStyle estiloOrigen = celdaOrigen.getCellStyle();
                        CellStyle estiloDestino = wbDestino.createCellStyle();
                   
                        // Copiar propiedades de estilo de la celda manualmente
                        copiarColorFondo(wbOrigen,wbDestino, estiloOrigen, estiloDestino);
                        celdaDestino.setCellStyle(estiloDestino);
                    }
                }
            }
        }

    }

    private static void copiarColorFondo(Workbook wbOrigen, Workbook wbDestino, CellStyle estiloOrigen, CellStyle estiloDestino) {
        // Copiar color de fondo y primer plano
        estiloDestino.setFillForegroundColor(estiloOrigen.getFillForegroundColor());
        estiloDestino.setFillPattern(estiloOrigen.getFillPattern());

        // Copiar fuentes
        Font fuenteOrigen = wbOrigen.getFontAt(estiloOrigen.getFontIndex());  
        Font fuenteDestino = wbDestino.createFont();
        copiarBold(fuenteOrigen, fuenteDestino);
        estiloDestino.setFont(fuenteDestino);
    }

    // Función para copiar las propiedades de la fuente
    private static void copiarBold(Font fuenteOrigen, Font fuenteDestino) {
    fuenteDestino.setBold(fuenteOrigen.getBold());
    fuenteDestino.setUnderline(fuenteOrigen.getUnderline());
    }

    public static void main(String[] args) {
        String inputFilePath = "O:/aa/test-lunes-jueves.xlsx";
        String outputFilePath = "O:/aa/result.xlsx";
        try {
            lunes_jueves(inputFilePath, outputFilePath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
