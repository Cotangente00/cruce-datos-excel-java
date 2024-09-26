package procesamiento_hojas.manipular_INFORME_SOLICITUDES;

import org.apache.poi.ss.usermodel.*;
import java.io.IOException;

public class borrar_columnas_restantes {
    public static void borrarColumnasRestantes(Workbook wb) throws IOException {
    
        Sheet ws = wb.getSheetAt(0); // Obtener la primera hoja

        // Índices de las columnas a eliminar (empiezan desde 0: A=0, B=1, C=2, etc.)
        int[] columnasAEliminar = {14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25};

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
    }
}
