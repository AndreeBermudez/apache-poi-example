package com.excel.prueba;

import org.apache.commons.codec.DecoderException;
import org.apache.commons.codec.binary.Hex;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GeneradorFuentes {
    public static class Builder{

        private String nombreFuente;
        private boolean conNegrita;
        private boolean conItalica;
        private short tamanoFuente;
        private short colorDefecto;
        private XSSFColor colorPersonalizado;
        private FontUnderline tipoUnderline;

        public Builder setNombreFuente(String nombreFuente) {
            this.nombreFuente = nombreFuente;
            return this;
        }

        public Builder setConNegrita(boolean conNegrita){
            this.conNegrita = conNegrita;
            return this;
        }

        public Builder setConItalica(boolean conItalica){
            this.conItalica = conItalica;
            return this;
        }

        public Builder setTamanoFuente(short tamanoFuente){
            this.tamanoFuente = tamanoFuente;
            return this;
        }

        public Builder setColorDefecto(short colorDefecto){
            this.colorDefecto = colorDefecto;
            return this;
        }

        public Builder setColorPersonalizado(String colorPersonalizado){
            try {
                byte[] rgb = Hex.decodeHex(colorPersonalizado);
                this.colorPersonalizado = new XSSFColor(rgb);
                return this;
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        }

        public Builder setTipoUnderline(FontUnderline tipoUnderline){
            this.tipoUnderline = tipoUnderline;
            return this;
        }

        public XSSFFont build(XSSFWorkbook libro){
            XSSFFont fuente = libro.createFont();
            if (nombreFuente != null) fuente.setFontName(nombreFuente);
            if (conNegrita) fuente.setBold(true);
            if (conItalica) fuente.setItalic(true);
            if (tamanoFuente != 0) fuente.setFontHeightInPoints(tamanoFuente);
            if (colorDefecto != 0) fuente.setColor(colorDefecto);
            if (colorPersonalizado != null) fuente.setColor(colorPersonalizado);
            if (tipoUnderline != null) fuente.setUnderline(tipoUnderline);
            return fuente;
        }

    }
}
