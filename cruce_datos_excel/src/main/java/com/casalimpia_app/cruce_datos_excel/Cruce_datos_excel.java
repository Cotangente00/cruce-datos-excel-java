/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 */

package com.casalimpia_app.cruce_datos_excel;

import com.casalimpia_app.cruce_datos_excel.model.FileProcessor;
import com.casalimpia_app.cruce_datos_excel.procesamiento_hojas.maquillaje.ventana_emergente_informativa;
import javafx.application.Application;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.layout.VBox;
import javafx.geometry.Pos;

import java.io.File;

/**
 *
 * @author jcavilaa
 */
public class Cruce_datos_excel extends Application {

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("Cruce de Datos Excel");

        // Crear el botón para seleccionar el archivo Excel
        Button selectFileButton = new Button("Seleccionar archivo Excel");
        selectFileButton.setOnAction(e -> {
            try {
                FileProcessor.processExcelFile(primaryStage);
                String detallesCambios = "Se realizaron los siguientes cambios:\n"
                    + "• Cambios en la hoja \"INFORME SOLICITUDES\":\n"
                    + "- Se eliminaron las filas 1, 2, 3, y 4.\n"
                    + "- Columnas eliminadas: Total servicios, Tipo, Turno partido, Jornada fija, Concepto: novedad y ausencias, Concepto: novedad control empleados, CC experta cambios, Experta cambio, Notificación SMS,  ";
                // Mostrar la ventana emergente con el mensaje de los cambios
                ventana_emergente_informativa.mostrarVentanaInformativa(detallesCambios);
                
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        });

        // Configurar el layout de la ventana
        VBox vbox = new VBox(10, selectFileButton);
        vbox.setAlignment(Pos.CENTER);
        Scene scene = new Scene(vbox, 420, 80);

        primaryStage.setScene(scene);
        primaryStage.show();
    }

    public static void main(String[] args) {
        launch(args);
    }
}