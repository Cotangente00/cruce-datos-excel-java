/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.cruce_datos_excel.procesamiento_hojas.maquillaje;

import javafx.application.Application;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.ButtonType;
import javafx.stage.Stage;

/**
 *
 * @author jcavilaa
 */
public class ventana_emergente_informativa {
    public static void mostrarVentanaInformativa(String mensaje) {
        Alert alert = new Alert(AlertType.INFORMATION); // Tipo de alerta: Informativa
        alert.setTitle("Cambios realizados");
        alert.setHeaderText("Los cambios se han completado con éxito");
        alert.setContentText(mensaje);

        // Mostrar la ventana y esperar a que el usuario presione "Aceptar"
        alert.showAndWait().ifPresent(response -> {
            if (response == ButtonType.OK) {
                System.out.println("El usuario ha leído el mensaje y ha presionado Aceptar.");
            }
        });
    }
}
