package main;
import manipular_INFORME_SOLICITUDES.*;
import manipular_Hoja1.*;
import maquillaje.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Main {
    
    public static void main(String[] args) throws Exception {
        String inputFilePath = "O:/aa/test-lunes-jueves.xlsx";
        String outputFilePath = "O:/aa/result.xlsx";
        Workbook wb;
        if (inputFilePath.endsWith(".xlsx")) {
            wb = new XSSFWorkbook(new FileInputStream(inputFilePath));
        } else if (inputFilePath.endsWith(".xls")) {
            wb = new HSSFWorkbook(new FileInputStream(inputFilePath));
        } else {
            throw new IllegalArgumentException("Formato de archivo no soportado");
        }
        try {
            //funciones para INFORME SOLICITUDES
            eliminar_filas.delete_filas(wb);
            eliminar_columnas.eliminarColumnas(wb);
            borrar_columnas_restantes.borrarColumnasRestantes(wb);
            filtrar_ciudades.filtrarCiudades(wb);
            date_format.formatearFechas(wb);
            int_format.convertirTextoANumero(wb);
            novedades_expertas.resaltarNovedad(wb);
            //funciones para Hoja1
            find_table.encontrar_tabla(wb);
            concatenar_nombres_apellidos.concatenacion(wb);
            buscarV_nombres_cedulas.simulateBUSCARV(wb);
            buscarV_nombres.simulateBUSCARVHoja1(wb);
            no_service_copy_paste.copiarFilas(wb);
            delete_celdas_vacias_H.limpiar_caracteres_invisibles(wb);
            //Funciones para los estilos de ambas hojas
            horizontal_column_size.ajustarAnchoColumnas(wb);
            order_alphabetic_INFORME_SOLICITUDES.reorganizeExcel_INFORME_SOLICITUDES(wb);
            order_alphabetic_Hoja1.reorganizeExcel_Hoja1(wb);
            delete_image.copiarContenidoHoja(wb);
            System.out.println("Archivo procesado exitosamente.");

            // Escribir los cambios en un archivo nuevo
            FileOutputStream fileOutputStream = new FileOutputStream(outputFilePath);
            
            wb.write(fileOutputStream);
            // Cerrar recursos
            fileOutputStream.close();
            wb.close();

        } catch (IOException e) {
            System.out.println("Ocurrió un error al procesar el archivo: " + e.getMessage());
        }
    }
}
