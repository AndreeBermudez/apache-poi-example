package com.excel;

import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.OutputStream;

public class EstilosExcel {
    public static void main(String[] args) {

        XSSFWorkbook libro = new XSSFWorkbook();
        XSSFSheet hoja = libro.createSheet();
        XSSFRow fila = hoja.createRow(0);
        XSSFCell columna = fila.createCell(0);
        XSSFCellStyle estiloCelda =libro.createCellStyle();
        columna.setCellValue("Hola Mundo");

        try {
            OutputStream file = new FileOutputStream("EstilosExcel.xlsx");
            libro.write(file);

            file.close();
            libro.close();
        }catch(Exception e){
            System.out.println("Error: " + e.getMessage());
        }

    }
}
