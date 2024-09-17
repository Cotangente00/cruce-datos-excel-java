package main;
import manipular_INFORME_SOLICITUDES.*;
import manipular_Hoja1.*;
import maquillaje.*;
import java.io.IOException;


public class Main {
    
    public static void main(String[] args) throws Exception {
        String rutaArchivo = "O:/aa/result.xlsx";
        eliminar_filas.main(args);
        try {
            eliminar_columnas.eliminarColumnas(rutaArchivo);
            borrar_columnas_restantes.borrarColumnasRestantes(rutaArchivo);
            filtrar_ciudades.filtrarCiudades(rutaArchivo);
            date_format.formatearFechas(rutaArchivo);
            int_format.convertirTextoANumero(rutaArchivo);
            novedades_expertas.resaltarNovedad(rutaArchivo);
            find_table.encontrar_tabla(rutaArchivo);
            concatenar_nombres_apellidos.concatenacion(rutaArchivo);
            buscarV_nombres_cedulas.simulateBUSCARV(rutaArchivo);
            buscarV_nombres.simulateBUSCARVHoja1(rutaArchivo);
            no_service_copy_paste.copiarFilas(rutaArchivo);
            delete_celdas_vacias_H.limpiar_caracteres_invisibles(rutaArchivo);
            horizontal_column_size.ajustarAnchoColumnas(rutaArchivo);
            order_alphabetic_INFORME_SOLICITUDES.reorganizeExcel_INFORME_SOLICITUDES(rutaArchivo);
            order_alphabetic_Hoja1.reorganizeExcel_Hoja1(rutaArchivo);
            delete_image.copiarContenidoHoja(rutaArchivo);
            System.out.println("Columnas eliminadas correctamente.");
        } catch (IOException e) {
            System.out.println("Ocurrió un error al procesar el archivo: " + e.getMessage());
        }
    }
}
