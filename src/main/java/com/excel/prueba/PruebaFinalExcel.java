package com.excel.prueba;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

public class PruebaFinalExcel {

    public static void main(String[] args) {
        List<Cliente> listadoClientes = obtenerListado();
        Field[] campos = Cliente.class.getDeclaredFields();

        XSSFWorkbook libro = new XSSFWorkbook();
        XSSFSheet hoja = libro.createSheet("Clientes");

        XSSFFont fuenteTitulo = new GeneradorFuentes.Builder()
                .setNombreFuente("Berlin Sans FB Demi")
                .setTamanoFuente((short) 14)
                .setConNegrita(true)
                .setTipoUnderline(FontUnderline.SINGLE)
                .build(libro);

        XSSFFont fuenteContenido = new GeneradorFuentes.Builder()
                .setNombreFuente("Calibri")
                .setTamanoFuente((short) 12)
                .setConItalica(true)
                .build(libro);

        XSSFCellStyle estiloTitulo = new GeneradorEstilos.Builder()
                .setColorPersonalizado("4f8cc9")
                .setTipoPatron(FillPatternType.SOLID_FOREGROUND)
                .setAlineacionHorizonal(HorizontalAlignment.CENTER)
                .setBordeArriba(BorderStyle.THIN)
                .setBordeAbajo(BorderStyle.THIN)
                .setBordeDerecho(BorderStyle.THIN)
                .setBordeIzquierdo(BorderStyle.THIN)
                .setFuente(fuenteTitulo)
                .build(libro);

        XSSFCellStyle estilosContenido = new GeneradorEstilos.Builder()
                .setColorPersonalizado("98cbff")
                .setTipoPatron(FillPatternType.SOLID_FOREGROUND)
                .setAlineacionHorizonal(HorizontalAlignment.CENTER)
                .setBordeArriba(BorderStyle.THIN)
                .setBordeAbajo(BorderStyle.THIN)
                .setBordeDerecho(BorderStyle.THIN)
                .setBordeIzquierdo(BorderStyle.THIN)
                .setFuente(fuenteContenido)
                .build(libro);

        XSSFCellStyle estilosFecha = new GeneradorEstilos.Builder()
                .setColorPersonalizado("98cbff")
                .setTipoPatron(FillPatternType.SOLID_FOREGROUND)
                .setAlineacionHorizonal(HorizontalAlignment.CENTER)
                .setBordeArriba(BorderStyle.THIN)
                .setBordeAbajo(BorderStyle.THIN)
                .setBordeDerecho(BorderStyle.THIN)
                .setBordeIzquierdo(BorderStyle.THIN)
                .setFormato("dd/MM/yyyy")
                .setFuente(fuenteContenido)
                .build(libro);

        XSSFRow fila = null;
        XSSFCell celda = null;

        for (int i = 0; i < listadoClientes.size(); i++) {
            //Generamos la cabecera
            if(i==0){
                fila=hoja.createRow(i);
                for (int j = 0; j < campos.length; j++) {
                    celda=fila.createCell(j);
                    celda.setCellValue(campos[j].getName().toUpperCase());
                    celda.setCellStyle(estiloTitulo);
                }
            }

            Cliente cliente = listadoClientes.get(i);
            List<Object> atributos = cliente.obtenerAtributos();

            fila = hoja.createRow(i+1);

            for (int a = 0; a < atributos.size(); a++) {
                celda = fila.createCell(a);

                if (atributos.get(a) instanceof  Long){
                    celda.setCellValue((Long)atributos.get(a));
                    celda.setCellStyle(estilosContenido);
                }
                if (atributos.get(a) instanceof  String){
                    celda.setCellValue((String)atributos.get(a));
                    celda.setCellStyle(estilosContenido);
                }
                if (atributos.get(a) instanceof  LocalDate) {
                    celda.setCellValue((LocalDate) atributos.get(a));
                    celda.setCellStyle(estilosFecha);
                }
                hoja.autoSizeColumn(a);
            }

        }

        try {
            OutputStream output = new FileOutputStream("Clientes.xlsx");
            libro.write(output);

            libro.close();
            output.close();
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException("Error al crear el documento");
        }
    }

    public static List<Cliente> obtenerListado(){
        List<Cliente> listadoClientes = new ArrayList<>();
        listadoClientes.add(new Cliente(1L, "Santiago", "Perez", "12345", "santy@mail.com", LocalDate.of(1998, 11, 14)));
        listadoClientes.add(new Cliente(2L, "Anyi", "Hoyos", "756743", "anyi@mail.com", LocalDate.of(1999, 10, 06)));
        listadoClientes.add(new Cliente(3L, "Andrea", "Calle", "64574", "andre@mail.com", LocalDate.of(2000, 8, 02)));
        listadoClientes.add(new Cliente(4L, "Daniel", "Posada", "47383", "daniel@mail.com", LocalDate.of(1996, 11, 20)));
        listadoClientes.add(new Cliente(5L, "Freddy", "Socha", "678473", "fredy@mail.com", LocalDate.of(1990, 03, 25)));
        listadoClientes.add(new Cliente(6L, "Alejandra", "Posada", "16273", "aleja@mail.com", LocalDate.of(1991, 12, 05)));
        listadoClientes.add(new Cliente(7L, "Daniel", "Manrique", "474863", "danielManrique@mail.com", LocalDate.of(1980, 11, 25)));

        return listadoClientes;
    }
}
