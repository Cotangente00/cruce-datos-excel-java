package main;
import manipular_INFORME_SOLICITUDES.*;
import manipular_Hoja1.*;
import java.io.IOException;


public class Main {
    
    public static void main(String[] args) {
        String rutaArchivo = "O:/programa/cruce-datos-excel-java/result.xlsx";
        eliminar_filas.main(args);
        try {
            eliminar_columnas.eliminarColumnas(rutaArchivo);
            borrar_columnas_restantes.borrarColumnasRestantes(rutaArchivo);
            filtrar_ciudades.filtrarCiudades(rutaArchivo);
            date_format.formatearFechas(rutaArchivo);
            int_format.convertirTextoANumero(rutaArchivo);
            novedades_expertas.resaltarNovedad(rutaArchivo);
            find_table.encontrar_tabla(rutaArchivo);
            System.out.println("Columnas eliminadas correctamente.");
        } catch (IOException e) {
            System.out.println("Ocurrió un error al procesar el archivo: " + e.getMessage());
        }
    }
}
