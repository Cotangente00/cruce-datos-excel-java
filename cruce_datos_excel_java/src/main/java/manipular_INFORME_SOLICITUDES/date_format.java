package manipular_INFORME_SOLICITUDES;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.DateUtil;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

public class date_format {
    public static void formatearFechas(String inputFilePath) throws IOException {
        // Cargar el archivo Excel
        FileInputStream fileInputStream = new FileInputStream(inputFilePath);
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);

        // Crear un estilo de celda para formato de fecha DD/MM/YYYY
        CellStyle dateCellStyle = workbook.createCellStyle();
        CreationHelper createHelper = workbook.getCreationHelper();
        dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("DD/MM/YYYY"));

        // Iterar sobre las filas para convertir y formatear las fechas en la columna D (índice 3)
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) { // Inicia en 1 para saltar el encabezado
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                Cell cellFecha = row.getCell(3); // Columna D es el índice 3 (0-indexed)

                if (cellFecha != null && cellFecha.getCellType() == CellType.NUMERIC) {
                    double numericValue = cellFecha.getNumericCellValue();

                    // Verificar si el número puede ser interpretado como una fecha
                    if (DateUtil.isValidExcelDate(numericValue)) {
                        // Convertir el valor numérico a una fecha
                        Date date = DateUtil.getJavaDate(numericValue);

                        // Cambiar el tipo de celda a fecha y aplicar el formato
                        cellFecha.setCellValue(date);
                        cellFecha.setCellStyle(dateCellStyle);
                    }
                }
            }
        }


        // Guardar los cambios en un nuevo archivo
        FileOutputStream fileOutputStream = new FileOutputStream(inputFilePath);
        workbook.write(fileOutputStream);

        // Cerrar recursos
        fileOutputStream.close();
        workbook.close();
        fileInputStream.close();

        System.out.println("Proceso completado. Fechas formateadas y archivo guardado en: " + inputFilePath);    
    }
}