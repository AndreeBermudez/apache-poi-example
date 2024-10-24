package com.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.OutputStream;

public class PrincipalExcel {
    
    public static void main(String[] args) {

        //1. Crear un libro de trabajo en Excel
        Workbook libro = new XSSFWorkbook();

        //2. Crear una hoja en el libro
        Sheet hoja = libro.createSheet("Reporte de Ventas");

        //3. Crear las filas
        Row cabecera = hoja.createRow(2);

        //4. Crear las columnas
        Cell nombre = cabecera.createCell(1);
        Cell edad = cabecera.createCell(2);
        Cell ciudad = cabecera.createCell(3);
        nombre.setCellValue("Nombre");
        edad.setCellValue("Edad");
        ciudad.setCellValue("Ciudad");


        Row fila1 = hoja.createRow(3);
        Cell col1 = fila1.createCell(1);
        Cell col2 = fila1.createCell(2);
        Cell col3 = fila1.createCell(3);
        col1.setCellValue("Andree");
        col2.setCellValue(23);
        col3.setCellValue("Chimbote");

        Row fila2 = hoja.createRow(4);
        Cell col11 = fila2.createCell(1);
        Cell col21 = fila2.createCell(2);
        Cell col31 = fila2.createCell(3);
        col11.setCellValue("Harold");
        col21.setCellValue(23);
        col31.setCellValue("Chimbote");

        try {
            OutputStream output = new FileOutputStream("ArchivoExcel.xlsx");
            libro.write(output);
            //Liberar recursos luego de exportar el archivo
            libro.close();
            output.close();
        }catch (Exception e){
            e.printStackTrace();
        }
    }
}

//        for (int i=2; i<5;i++){
//            Row filas = hoja.createRow(i);
//            for (int j=1; j<4;j++){
//                Cell celdas = filas.createCell(j);
//                celdas.setCellValue("Holis");
//            }
//        }
