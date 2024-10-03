/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 */

package com.casalimpia_app.cruce_datos_excel;

import com.casalimpia_app.cruce_datos_excel.model.FileProcessor;
import javafx.application.Application;
import javafx.stage.Stage;

/**
 *
 * @author jcavilaa
 */
public class Cruce_datos_excel extends Application {

    @Override
    public void start(Stage primaryStage) throws Exception {
        primaryStage.setWidth(420);
        primaryStage.setHeight(80);
        primaryStage.setResizable(false);
        primaryStage.setTitle("Cruce de Datos Excel");

        // Llamar al método para procesar el archivo Excel
        primaryStage.show();
        FileProcessor.processExcelFile(primaryStage);
    }

    public static void main(String[] args) {
        launch(args);
    }
}