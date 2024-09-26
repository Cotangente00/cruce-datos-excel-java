package model;

//import org.apache.poi.ss.usermodel.*;
//import java.io.*;
import java.time.*;
import java.nio.file.*;

public class boton {
    public static void boton_excel(String inputFilePath, String outputFilePath) throws Exception {
        Path filePath = Paths.get(inputFilePath);
        Instant instant = Files.getLastModifiedTime(filePath).toInstant();
        LocalDate creationDate = instant.atZone(ZoneId.systemDefault()).toLocalDate();

        // Determinar el día de la semana
        DayOfWeek dayOfWeek = creationDate.getDayOfWeek();

        // Ejecutar la función correspondiente
        if (dayOfWeek == DayOfWeek.MONDAY || dayOfWeek == DayOfWeek.TUESDAY ||
                dayOfWeek == DayOfWeek.WEDNESDAY || dayOfWeek == DayOfWeek.THURSDAY) {
            funciones_lunes_jueves.lunes_jueves(inputFilePath, outputFilePath);
        } else if (dayOfWeek == DayOfWeek.FRIDAY || dayOfWeek == DayOfWeek.SATURDAY) {
            funciones_viernes_sabado.viernes_sabado(inputFilePath, outputFilePath);
        } else {
            System.out.println("La fecha de creación del archivo no es válida o cae en domingo.");
        }

    }   

    public static void main(String[] args) {
        String inputFilePath = "O:/aa/test-lunes-jueves.xlsx";
        String outputFilePath = "O:/aa/result.xlsx";
        try {
            boton_excel(inputFilePath, outputFilePath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
