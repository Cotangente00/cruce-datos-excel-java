package manipular_INFORME_SOLICITUDES;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class borrar_columnas_restantes {
    public static void borrarColumnasRestantes(String archivoExcel) throws IOException {
        // Cargar el archivo Excel
        FileInputStream archivo = new FileInputStream(archivoExcel);
        Workbook wb = new XSSFWorkbook(archivo);
        Sheet ws = wb.getSheetAt(0); // Obtener la primera hoja

        // Índices de las columnas a eliminar (empiezan desde 0: A=0, B=1, C=2, etc.)
        int[] columnasAEliminar = {13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25};

        // Recorrer todas las filas de la hoja
        for (Row fila : ws) {

            // Recorrer las columnas a eliminar en orden inverso para evitar problemas
            for (int i = columnasAEliminar.length - 1; i >= 0; i--) {
                int columna = columnasAEliminar[i];
                Cell celda = fila.getCell(columna);
                if (celda != null) {
                    fila.removeCell(celda);
                }       
            }
        }
    
        // Guardar los cambios en el archivo
        FileOutputStream archivoSalida = new FileOutputStream(archivoExcel);
        wb.write(archivoSalida);

        // Cerrar los archivos y liberar recursos
        archivoSalida.close();
        wb.close();
        archivo.close();
    }
}
