/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.cruce_datos_excel.model;

import com.casalimpia_app.cruce_datos_excel.Cruce_datos_excel;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import java.io.File;
import java.io.IOException;
import java.nio.file.*;
import java.time.DayOfWeek;
import java.time.Instant;
import java.time.LocalDate;
import java.time.ZoneId;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;


/**
 *
 * @author jcavilaa
 */
public class FileProcessor {

    public static void processExcelFile(Stage stage) throws IOException, EncryptedDocumentException, InvalidFormatException, Exception {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Seleccionar archivo Excel de entrada");

        // Filtro para archivos Excel
        fileChooser.getExtensionFilters().addAll(
            new FileChooser.ExtensionFilter("Archivos Excel", "*.xls", "*.xlsx")
        );

        // Abrir ventana para seleccionar archivo de entrada
        File inputFile = fileChooser.showOpenDialog(stage);
        if (inputFile == null) {
            System.out.println("No se seleccionó un archivo.");
            return;
        }
        String inputFilePath = inputFile.getAbsolutePath();
        Path filePath = Paths.get(inputFilePath); 
        Instant instant = Files.getLastModifiedTime(filePath).toInstant();
        LocalDate creationDate = instant.atZone(ZoneId.systemDefault()).toLocalDate();

        // Determinar el día de la semana
        DayOfWeek dayOfWeek = creationDate.getDayOfWeek();

        // Abrir ventana para seleccionar ubicación de guardado del archivo de salida
        fileChooser.setTitle("Guardar archivo Excel modificado");
        fileChooser.setInitialFileName("result.xlsx");  // Nombre por defecto para el archivo de salida
        File outputFile = fileChooser.showSaveDialog(stage);
        if (outputFile == null) {
            System.out.println("No se seleccionó una ubicación para guardar el archivo.");
            return;
        }
        String outputFilePath = outputFile.getAbsolutePath();

        // Ejecutar la función correspondiente según el día de la semana
        if (dayOfWeek != null) {
            switch (dayOfWeek) {
                case MONDAY:
                case TUESDAY:
                case WEDNESDAY:
                case THURSDAY:
                    processorLunes_Jueves.lunes_jueves(inputFilePath, outputFilePath);
                    break;
                case FRIDAY:
                case SATURDAY:
                    processorViernes_sabado.viernes_sabado(inputFilePath, outputFilePath);
                    break;
                default:
                    System.out.println("La fecha de creación del archivo no es válida o cae en domingo.");
                    break;
            }
        }
    }

    public static void main(String[] args) {
        javafx.application.Application.launch(Cruce_datos_excel.class, args);
    }


    /*
    public void saveProcessedFile(File outputFile) {
        try (FileOutputStream fos = new FileOutputStream(outputFile);
             Workbook workbook = new XSSFWorkbook()) {

            // Crear una nueva hoja o trabajar con los datos procesados
            Sheet sheet = workbook.createSheet("Procesado");

            // Guardar los cambios en un archivo nuevo
            workbook.write(fos);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }*/
}
