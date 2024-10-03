/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.cruce_datos_excel.procesamiento_hojas.maquillaje;

import java.io.IOException;

import org.apache.poi.ss.usermodel.*;

/**
 *
 * @author jcavilaa
 */
public class estilos_encabezados_xls {
    public static void estilos_encabezados(Workbook wb) throws IOException {
        Sheet ws = wb.getSheetAt(0); //Hoja INFORME_SOLICITUDES
        Row fila = ws.getRow(0); //Fila A

        // Crear un estilo con negrita y subrayado
        CellStyle estilo = wb.createCellStyle();
        Font fuente = wb.createFont();
        fuente.setBold(true);
        fuente.setUnderline(Font.U_SINGLE);
        estilo.setFont(fuente);
        
        //Aplicar los estilos por cada celda de la fila
        for (Cell celda : fila) {
            celda.setCellStyle(estilo);
        }
    }
}
