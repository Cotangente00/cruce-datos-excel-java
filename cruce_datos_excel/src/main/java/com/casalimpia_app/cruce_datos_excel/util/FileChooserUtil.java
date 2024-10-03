/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.cruce_datos_excel.util;

import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.io.File;

/**
 *
 * @author jcavilaa
 */
public class FileChooserUtil {

    public static File openFileChooser(Stage stage) {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Selecciona un archivo Excel");
        fileChooser.getExtensionFilters().addAll(
            new FileChooser.ExtensionFilter("Excel Files", "*.xls", "*.xlsx")
        );
        return fileChooser.showOpenDialog(stage);
    }

    public static File openSaveFileChooser(Stage stage) {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Guardar archivo procesado");
        fileChooser.getExtensionFilters().addAll(
            new FileChooser.ExtensionFilter("Excel Files", "*.xlsx")
        );
        return fileChooser.showSaveDialog(stage);
    }
}
