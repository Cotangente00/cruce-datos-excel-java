package main;
import manipular_INFORME_SOLICITUDES.*;
import manipular_Hoja1.*;
import maquillaje.*;
import java.io.IOException;


public class Main {
    
    public static void main(String[] args) throws Exception {
        String rutaArchivo = "O:/aa/test-lunes-jueves.xlsx";
        String rutaArchivoSalida = "O:/aa/result.xlsx";
        try {
            eliminar_filas.delete_filas(rutaArchivo, rutaArchivoSalida);
            eliminar_columnas.eliminarColumnas(rutaArchivoSalida);
            borrar_columnas_restantes.borrarColumnasRestantes(rutaArchivoSalida);
            filtrar_ciudades.filtrarCiudades(rutaArchivoSalida);
            date_format.formatearFechas(rutaArchivoSalida);
            int_format.convertirTextoANumero(rutaArchivoSalida);
            novedades_expertas.resaltarNovedad(rutaArchivoSalida);
            find_table.encontrar_tabla(rutaArchivoSalida);
            concatenar_nombres_apellidos.concatenacion(rutaArchivoSalida);
            buscarV_nombres_cedulas.simulateBUSCARV(rutaArchivoSalida);
            buscarV_nombres.simulateBUSCARVHoja1(rutaArchivoSalida);
            no_service_copy_paste.copiarFilas(rutaArchivoSalida);
            delete_celdas_vacias_H.limpiar_caracteres_invisibles(rutaArchivoSalida);
            horizontal_column_size.ajustarAnchoColumnas(rutaArchivoSalida);
            order_alphabetic_INFORME_SOLICITUDES.reorganizeExcel_INFORME_SOLICITUDES(rutaArchivoSalida);
            order_alphabetic_Hoja1.reorganizeExcel_Hoja1(rutaArchivoSalida);
            delete_image.copiarContenidoHoja(rutaArchivoSalida);
            System.out.println("Columnas eliminadas correctamente.");
        } catch (IOException e) {
            System.out.println("Ocurrió un error al procesar el archivo: " + e.getMessage());
        }
    }
}
