package com.classes.methods;

import com.app.app;
import com.aspose.cells.*;
import com.classes.connection.conexion;

import javafx.application.Platform;
import javafx.scene.control.Alert;
import javafx.stage.Stage;
import java.io.File;
import java.io.PrintWriter;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

public class generateXLS {
    String typeExcel = "";
    int codeAlert = 0;

    public void excelInformes(String month, String anio) {
        String mes = "" + month.charAt(0) + month.charAt(1);
        if (month.charAt(1) == ' ') {
            mes = "0" + month.charAt(0);
        }

        conexion sql = new conexion(); //CREA UNA NUEVA CONEXION CON LA BASE DE DATOS
        Connection con = sql.conectarSQL(); //LLAMA LA CONEXION

        List<String> listPorciones = new ArrayList<String>();

        try {
            PreparedStatement ps = con.prepareStatement("SELECT DISTINCT porcion FROM IMPRESION WHERE (f_cierre BETWEEN '"+anio+"-"+mes+"-01' AND '"+anio+"-"+mes+"-31') ORDER BY porcion");
            ResultSet rs = ps.executeQuery();
            while (rs.next()) {
                String porcion = rs.getString("porcion");
                listPorciones.add(porcion);
            }

        } catch (Exception e) {
            System.out.println(e);
        }

        String querySuspensiones = "";
        for (int i = 0; i < listPorciones.size(); i++) {
            querySuspensiones += "SELECT (SELECT ('"+listPorciones.get(i)+"')) porcion, " +
                    "(SELECT f_ejecutado FROM SyT WHERE (porcion = '"+listPorciones.get(i)+"' AND f_cierre BETWEEN '"+anio+"-"+mes+"-01' AND '"+anio+"-"+mes+"-31' AND tipo_solicitud = 'SUSPENSION')) f_dunning, " +
                    "(SELECT f_cierre FROM SyT WHERE (porcion = '"+listPorciones.get(i)+"' AND f_cierre BETWEEN '"+anio+"-"+mes+"-01' AND '"+anio+"-"+mes+"-31' AND tipo_solicitud = 'SUSPENSION')) f_cierre, " +
                    "(SELECT count(*) FROM SyT WHERE (porcion = '"+listPorciones.get(i)+"' AND f_cierre BETWEEN '"+anio+"-"+mes+"-01' AND '"+anio+"-"+mes+"-31' AND tipo_solicitud = 'SUSPENSION')) total_suspensiones, " +
                    "(SELECT sum(pagos) FROM SyT WHERE (porcion = '"+listPorciones.get(i)+"' AND f_cierre BETWEEN '"+anio+"-"+mes+"-01' AND '"+anio+"-"+mes+"-31' AND tipo_solicitud = 'SUSPENSION')) valor_total_cartera, " +
                    "(SELECT count(*) FROM SyT WHERE (porcion = '"+listPorciones.get(i)+"' AND f_cierre BETWEEN '"+anio+"-"+mes+"-01' AND '"+anio+"-"+mes+"-31' AND tipo_solicitud = 'SUSPENSION' AND resultado = 1)) efectivas, " +
                    "(SELECT count(*) FROM SyT WHERE (porcion = '"+listPorciones.get(i)+"' AND f_cierre BETWEEN '"+anio+"-"+mes+"-01' AND '"+anio+"-"+mes+"-31' AND tipo_solicitud = 'SUSPENSION' AND resultado = 3)) pagos, " +
                    "(SELECT count(*) FROM SyT WHERE (porcion = '"+listPorciones.get(i)+"' AND f_cierre BETWEEN '"+anio+"-"+mes+"-01' AND '"+anio+"-"+mes+"-31' AND tipo_solicitud = 'SUSPENSION' AND resultado = 34)) conserva_estado, " +
                    "(SELECT count(*) FROM SyT WHERE (porcion = '"+listPorciones.get(i)+"' AND f_cierre BETWEEN '"+anio+"-"+mes+"-01' AND '"+anio+"-"+mes+"-31' AND tipo_solicitud = 'SUSPENSION' AND resultado != 1  AND resultado != 3  AND resultado != 34)) otras_anomalias, " +
                    "(SELECT sum(pagos) FROM SyT WHERE (porcion = '"+listPorciones.get(i)+"' AND f_cierre BETWEEN '"+anio+"-"+mes+"-01' AND '"+anio+"-"+mes+"-31' AND tipo_solicitud = 'SUSPENSION' AND resultado = 1)) valor_cartera_efectiva, " +
                    "(SELECT sum(pagos) FROM IMPRESION WHERE (porcion = '"+listPorciones.get(i)+"' AND f_cierre BETWEEN '"+anio+"-"+mes+"-01' AND '"+anio+"-"+mes+"-31' AND tipo_solicitud = 'SUSPENSION')) valor_cartera_impresa, " +
                    "(SELECT sum(pagos) FROM EXCLUIDAS WHERE (porcion = '"+listPorciones.get(i)+"' AND f_cierre BETWEEN '"+anio+"-"+mes+"-01' AND '"+anio+"-"+mes+"-31' AND tipo_solicitud = 'SUSPENSION')) valor_cartera_excluida";
            if (i < (listPorciones.size()-1)) {
                querySuspensiones += " UNION ";
            }
        }

        String queryTaponamientos = "";
        for (int i = 0; i < listPorciones.size(); i++) {
            queryTaponamientos += "SELECT (SELECT ('"+listPorciones.get(i)+"')) porcion, " +
                    "(SELECT f_ejecutado FROM SyT WHERE (porcion = '"+listPorciones.get(i)+"' AND f_cierre BETWEEN '"+anio+"-"+mes+"-01' AND '"+anio+"-"+mes+"-31' AND tipo_solicitud = 'TAPONAMIENTO')) f_dunning, " +
                    "(SELECT f_cierre FROM SyT WHERE (porcion = '"+listPorciones.get(i)+"' AND f_cierre BETWEEN '"+anio+"-"+mes+"-01' AND '"+anio+"-"+mes+"-31' AND tipo_solicitud = 'TAPONAMIENTO')) f_cierre, " +
                    "(SELECT count(*) FROM SyT WHERE (porcion = '"+listPorciones.get(i)+"' AND f_cierre BETWEEN '"+anio+"-"+mes+"-01' AND '"+anio+"-"+mes+"-31' AND tipo_solicitud = 'TAPONAMIENTO')) total_suspensiones, " +
                    "(SELECT sum(pagos) FROM SyT WHERE (porcion = '"+listPorciones.get(i)+"' AND f_cierre BETWEEN '"+anio+"-"+mes+"-01' AND '"+anio+"-"+mes+"-31' AND tipo_solicitud = 'TAPONAMIENTO')) valor_total_cartera, " +
                    "(SELECT count(*) FROM SyT WHERE (porcion = '"+listPorciones.get(i)+"' AND f_cierre BETWEEN '"+anio+"-"+mes+"-01' AND '"+anio+"-"+mes+"-31' AND tipo_solicitud = 'TAPONAMIENTO' AND resultado = 1)) efectivas, " +
                    "(SELECT count(*) FROM SyT WHERE (porcion = '"+listPorciones.get(i)+"' AND f_cierre BETWEEN '"+anio+"-"+mes+"-01' AND '"+anio+"-"+mes+"-31' AND tipo_solicitud = 'TAPONAMIENTO' AND resultado = 3)) pagos, " +
                    "(SELECT count(*) FROM SyT WHERE (porcion = '"+listPorciones.get(i)+"' AND f_cierre BETWEEN '"+anio+"-"+mes+"-01' AND '"+anio+"-"+mes+"-31' AND tipo_solicitud = 'TAPONAMIENTO' AND resultado = 34)) conserva_estado, " +
                    "(SELECT count(*) FROM SyT WHERE (porcion = '"+listPorciones.get(i)+"' AND f_cierre BETWEEN '"+anio+"-"+mes+"-01' AND '"+anio+"-"+mes+"-31' AND tipo_solicitud = 'TAPONAMIENTO' AND resultado != 1  AND resultado != 3  AND resultado != 34)) otras_anomalias, " +
                    "(SELECT sum(pagos) FROM SyT WHERE (porcion = '"+listPorciones.get(i)+"' AND f_cierre BETWEEN '"+anio+"-"+mes+"-01' AND '"+anio+"-"+mes+"-31' AND tipo_solicitud = 'TAPONAMIENTO' AND resultado = 1)) valor_cartera_efectiva, " +
                    "(SELECT sum(pagos) FROM IMPRESION WHERE (porcion = '"+listPorciones.get(i)+"' AND f_cierre BETWEEN '"+anio+"-"+mes+"-01' AND '"+anio+"-"+mes+"-31' AND tipo_solicitud = 'TAPONAMIENTO')) valor_cartera_impresa, " +
                    "(SELECT sum(pagos) FROM EXCLUIDAS WHERE (porcion = '"+listPorciones.get(i)+"' AND f_cierre BETWEEN '"+anio+"-"+mes+"-01' AND '"+anio+"-"+mes+"-31' AND tipo_solicitud = 'TAPONAMIENTO')) valor_cartera_excluida";
            if (i < (listPorciones.size()-1)) {
                queryTaponamientos += " UNION ";
            }
        }

        List<List<String>> dataSuspensiones = new ArrayList<>();
        List<List<String>> dataTaponamientos = new ArrayList<>();
        for (int i = 1; i <= 14; i++) {
            dataSuspensiones.add(new ArrayList<>());
            dataTaponamientos.add(new ArrayList<>());
        }
        List<List<String>> dataReinstalaciones = new ArrayList<>();
        for (int i = 1; i <= 6; i++) {
            dataReinstalaciones.add(new ArrayList<>());
        }

        List<List<String>> dataPorcionXreinstalaciones = new ArrayList<>();
        for (int i = 1; i <= 2; i++) {
            dataPorcionXreinstalaciones.add(new ArrayList<>());
        }

        try {
            PreparedStatement ps = con.prepareStatement(querySuspensiones);
            ResultSet rs = ps.executeQuery();
            while (rs.next()) {
                dataSuspensiones.get(0).add(rs.getString("porcion"));
                dataSuspensiones.get(1).add(rs.getString("f_dunning"));
                dataSuspensiones.get(2).add(rs.getString("f_cierre"));
                dataSuspensiones.get(3).add("");
                dataSuspensiones.get(4).add(rs.getString("total_suspensiones"));
                dataSuspensiones.get(5).add(rs.getString("valor_total_cartera"));
                dataSuspensiones.get(6).add(rs.getString("efectivas"));
                dataSuspensiones.get(7).add(rs.getString("pagos"));
                dataSuspensiones.get(8).add(rs.getString("conserva_estado"));
                dataSuspensiones.get(9).add(rs.getString("otras_anomalias"));
                dataSuspensiones.get(10).add(rs.getString("valor_cartera_efectiva"));
                dataSuspensiones.get(11).add(rs.getString("valor_cartera_impresa"));
                dataSuspensiones.get(12).add("");
                dataSuspensiones.get(13).add(rs.getString("valor_cartera_excluida"));
            }

            ps = con.prepareStatement(queryTaponamientos);
            rs = ps.executeQuery();
            while (rs.next()) {
                dataTaponamientos.get(0).add(rs.getString("porcion"));
                dataTaponamientos.get(1).add(rs.getString("f_dunning"));
                dataTaponamientos.get(2).add(rs.getString("f_cierre"));
                dataTaponamientos.get(3).add("");
                dataTaponamientos.get(4).add(rs.getString("total_suspensiones"));
                dataTaponamientos.get(5).add(rs.getString("valor_total_cartera"));
                dataTaponamientos.get(6).add(rs.getString("efectivas"));
                dataTaponamientos.get(7).add(rs.getString("pagos"));
                dataTaponamientos.get(8).add(rs.getString("conserva_estado"));
                dataTaponamientos.get(9).add(rs.getString("otras_anomalias"));
                dataTaponamientos.get(10).add(rs.getString("valor_cartera_efectiva"));
                dataTaponamientos.get(11).add(rs.getString("valor_cartera_impresa"));
                dataTaponamientos.get(12).add("");
                dataTaponamientos.get(13).add(rs.getString("valor_cartera_excluida"));
            }

            ps = con.prepareStatement("SELECT DISTINCT (fecha) reinstalacion, (f_cierre) cargue FROM REINSTALACIONES WHERE (f_cierre BETWEEN '"+anio+"-"+mes+"-01' AND '"+anio+"-"+mes+"-31')");
            rs = ps.executeQuery();
            while (rs.next()) {
                dataReinstalaciones.get(0).add(rs.getString("reinstalacion"));
                dataReinstalaciones.get(1).add(rs.getString("cargue"));
            }

            String queryReinstalaciones = "";
            for (int i = 0; i < dataReinstalaciones.get(0).size(); i++) {
                queryReinstalaciones += "SELECT (SELECT count() FROM REINSTALACIONES WHERE (fecha = '"+dataReinstalaciones.get(0).get(i)+"' AND f_cierre = '"+dataReinstalaciones.get(1).get(i)+"')) total, (SELECT count() FROM REINSTALACIONES WHERE (fecha = '"+dataReinstalaciones.get(0).get(i)+"' AND f_cierre = '"+dataReinstalaciones.get(1).get(i)+"' AND resultado = 1)) efectivas, (SELECT count() FROM REINSTALACIONES WHERE (fecha = '"+dataReinstalaciones.get(0).get(i)+"' AND f_cierre = '"+dataReinstalaciones.get(1).get(i)+"' AND resultado != 1)) inefectivas";

                if (i < (dataReinstalaciones.get(0).size()-1)) {
                    queryReinstalaciones += " UNION ";
                }
            }

            ps = con.prepareStatement(queryReinstalaciones);
            rs = ps.executeQuery();
            while (rs.next()) {
                dataReinstalaciones.get(2).add("");
                dataReinstalaciones.get(3).add(rs.getString("total"));
                dataReinstalaciones.get(4).add(rs.getString("efectivas"));
                dataReinstalaciones.get(5).add(rs.getString("inefectivas"));
            }

            String queryPorcionXreinstalacion = "";
            for (char cell = 'A'; cell <= 'Z'; cell++) {
                if (cell == 'I' || cell == 'O' || cell == 'Y') {
                    cell++;
                }
                queryPorcionXreinstalacion += "SELECT ('"+ cell +"4') porcion, (count(*)) total FROM REINSTALACIONES WHERE (porcion = '"+cell+"4' AND f_cierre BETWEEN '"+anio+"-"+mes+"-01' AND '"+anio+"-"+mes+"-31') ";

                if (cell != 'Z') {
                    queryPorcionXreinstalacion += " UNION ";
                }
            }

            ps = con.prepareStatement(queryPorcionXreinstalacion);
            rs = ps.executeQuery();
            while (rs.next()) {
                dataPorcionXreinstalaciones.get(0).add(rs.getString("porcion"));
                dataPorcionXreinstalaciones.get(1).add(rs.getString("total"));
            }

            con.close();

            try {
                Workbook wb = new Workbook();
                String fileMes = "";
                if (month.charAt(1) == ' ') {
                    fileMes = month.substring(4);
                } else {
                    fileMes = month.substring(5);
                }

                Worksheet wsInforme = wb.getWorksheets().get(0);

                Cells cells;

                cells = wsInforme.getCells();
                //TAMAÑO DE LAS CELDAS
                cells.setColumnWidth(0, 1); //A
                cells.setColumnWidth(1, 1); //B
                cells.setColumnWidth(2, 8.86); //C
                cells.setColumnWidth(3, 15.14); //D
                cells.setColumnWidth(4, 13.71); //E
                cells.setColumnWidth(5, 10.43); //F
                cells.setColumnWidth(6, 17); //G
                cells.setColumnWidth(7, 19.57); //H
                cells.setColumnWidth(8, 10.14); //I
                cells.setColumnWidth(9, 7.86); //J
                cells.setColumnWidth(10, 10.86); //K
                cells.setColumnWidth(11, 11.57); //L
                cells.setColumnWidth(12, 21.14); //M
                cells.setColumnWidth(13, 21); //N
                cells.setColumnWidth(14, 13.86); //O
                cells.setColumnWidth(15, 21.71); //P
                cells.setColumnWidth(16, 1); //Q
                cells.setColumnWidth(17, 1); //R

                Style style;
                StyleFlag flag = new StyleFlag();
                Range range;

                style = wb.createStyle();
                flag.setBorders(true);
                flag.setAlignments(true);
                flag.setCellShading(true);
                flag.setFont(true);

                //TABLA SUSPENSION
                wsInforme.getCells().get("C5").setValue("SUSPENSIONES");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,192,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("C5:P5");
                range.applyStyle(style, flag);
                wsInforme.getCells().merge(4,2,1,14);
                style = new Style();
                //PORCION
                wsInforme.getCells().get("C6").setValue("PORCION");
                style.getFont().setBold(true);
                style.getFont().setColor(Color.getWhite());
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("C6:C7");
                range.applyStyle(style, flag);
                wsInforme.getCells().merge(5,2,2,1);
                style = new Style();
                //DUNNING
                wsInforme.getCells().get("D6").setValue("FECHA DUNNING");
                style.getFont().setBold(true);
                style.getFont().setColor(Color.getWhite());
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("D6:D7");
                range.applyStyle(style, flag);
                wsInforme.getCells().merge(5,3,2,1);
                style = new Style();
                //CIERRE
                wsInforme.getCells().get("E6").setValue("FECHA CIERRE");
                style.getFont().setBold(true);
                style.getFont().setColor(Color.getWhite());
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("E6:E7");
                range.applyStyle(style, flag);
                wsInforme.getCells().merge(5,4,2,1);
                style = new Style();
                //PROMEDIO
                wsInforme.getCells().get("F6").setValue("PROMEDIO\n(días)");
                style.getFont().setBold(true);
                style.getFont().setColor(Color.getWhite());
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setHorizontalAlignment(TextAlignmentType.DISTRIBUTED);
                range = wsInforme.getCells().createRange("F6:F7");
                range.applyStyle(style, flag);
                wsInforme.getCells().merge(5,5,2,1);
                style = new Style();
                //TOTAL SUSPENSIONES
                wsInforme.getCells().get("G6").setValue("TOTAL\nSUSPENSIONES");
                style.getFont().setBold(true);
                style.getFont().setColor(Color.getWhite());
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setHorizontalAlignment(TextAlignmentType.DISTRIBUTED);
                range = wsInforme.getCells().createRange("G6:G7");
                range.applyStyle(style, flag);
                wsInforme.getCells().merge(5,6,2,1);
                style = new Style();
                //VALOR TOTAL CARTERA
                wsInforme.getCells().get("H6").setValue("CARTERA TOTAL");
                style.getFont().setBold(true);
                style.getFont().setColor(Color.getWhite());
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                range = wsInforme.getCells().createRange("H6:H7");
                range.applyStyle(style, flag);
                wsInforme.getCells().merge(5,7,2,1);
                style = new Style();
                //RESULTADO
                wsInforme.getCells().get("I6").setValue("RESULTADO");
                style.getFont().setBold(true);
                style.getFont().setColor(Color.getWhite());
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("I6:L6");
                range.applyStyle(style, flag);
                wsInforme.getCells().merge(5,8,1,4);
                style = new Style();
                //R -> EFECTIVAS
                wsInforme.getCells().get("I7").setValue("EFECTIVAS");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,230,153));
                style.setPattern(BackgroundType.SOLID);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.setVerticalAlignment(TextAlignmentType.CENTER);
                wsInforme.getCells().get("I7").setStyle(style);
                style = new Style();
                //R -> PAGOS
                wsInforme.getCells().get("J7").setValue("PAGOS");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,230,153));
                style.setPattern(BackgroundType.SOLID);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.setVerticalAlignment(TextAlignmentType.CENTER);
                wsInforme.getCells().get("J7").setStyle(style);
                style = new Style();
                //R -> PAGOS
                wsInforme.getCells().get("K7").setValue("CONSERVA ESTADO");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,230,153));
                style.setPattern(BackgroundType.SOLID);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setHorizontalAlignment(TextAlignmentType.DISTRIBUTED);
                style.setVerticalAlignment(TextAlignmentType.CENTER);
                wsInforme.getCells().get("K7").setStyle(style);
                style = new Style();
                //R -> OTRAS ANOMALIAS
                wsInforme.getCells().get("L7").setValue("OTRAS ANOMALIAS");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,230,153));
                style.setPattern(BackgroundType.SOLID);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setHorizontalAlignment(TextAlignmentType.DISTRIBUTED);
                style.setVerticalAlignment(TextAlignmentType.CENTER);
                wsInforme.getCells().get("L7").setStyle(style);
                style = new Style();
                //VALOR CARTERA EFECTIVA
                wsInforme.getCells().get("M6").setValue("VALOR\nCARTERA\nEFECTIVA");
                style.getFont().setBold(true);
                style.getFont().setColor(Color.getWhite());
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setHorizontalAlignment(TextAlignmentType.DISTRIBUTED);
                style.setVerticalAlignment(TextAlignmentType.CENTER);
                range = wsInforme.getCells().createRange("M6:M7");
                range.applyStyle(style, flag);
                wsInforme.getCells().merge(5,12,2,1);
                style = new Style();
                //VALOR CARTERA ENVIADA A TERRENO
                wsInforme.getCells().get("N6").setValue("CARTERA\nENVIADA\nTERRENO");
                style.getFont().setBold(true);
                style.getFont().setColor(Color.getWhite());
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setHorizontalAlignment(TextAlignmentType.DISTRIBUTED);
                style.setVerticalAlignment(TextAlignmentType.CENTER);
                range = wsInforme.getCells().createRange("N6:N7");
                range.applyStyle(style, flag);
                wsInforme.getCells().merge(5,13,2,1);
                style = new Style();
                //PORCENTAJE CARTERA SUSPENDIDA
                wsInforme.getCells().get("O6").setValue("PORCENTAJE\nCARTERA\nSUSPENDIDA");
                style.getFont().setBold(true);
                style.getFont().setColor(Color.getWhite());
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setHorizontalAlignment(TextAlignmentType.DISTRIBUTED);
                style.setVerticalAlignment(TextAlignmentType.CENTER);
                range = wsInforme.getCells().createRange("O6:O7");
                range.applyStyle(style, flag);
                wsInforme.getCells().merge(5,14,2,1);
                style = new Style();
                //VALOR CARTERA EXCLUIDA
                wsInforme.getCells().get("P6").setValue("CARTERA\nEXCLUIDA");
                style.getFont().setBold(true);
                style.getFont().setColor(Color.getWhite());
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setHorizontalAlignment(TextAlignmentType.DISTRIBUTED);
                style.setVerticalAlignment(TextAlignmentType.CENTER);
                range = wsInforme.getCells().createRange("P6:P7");
                range.applyStyle(style, flag);
                wsInforme.getCells().merge(5,15,2,1);
                style = new Style();

                //DATA
                char c;
                int list = 0;
                int value = 0;

                Calendar dunning = Calendar.getInstance();
                Calendar cierre = Calendar.getInstance();

                boolean f_error = false;
                for (c = 'C'; c <= 'P'; c++) {
                    style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                    style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                    style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                    style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                    style.setHorizontalAlignment(TextAlignmentType.CENTER);

                    if (value == 0) {
                        wsInforme.getCells().get("" + c + "" + (8 + list)).setValue(dataSuspensiones.get(value).get(list));
                    } else if (value == 1) {
                        String fecha = dataSuspensiones.get(1).get(list);
                        if (fecha != null) {
                            f_error = false;
                            DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
                            Date date = dateFormat.parse(fecha);

                            dunning.setTime(date);

                            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("d/MM/yyyy");
                            fecha = simpleDateFormat.format(date);
                            wsInforme.getCells().get("" + c + "" + (8 + list)).setValue(fecha);
                        } else {
                            f_error = true;
                            wsInforme.getCells().get("" + c + "" + (8 + list)).setValue("");
                        }
                    } else if (value == 2) {
                        String fecha = dataSuspensiones.get(2).get(list);
                        if (fecha != null) {
                            f_error = false;
                            DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
                            Date date = dateFormat.parse(fecha);
                            cierre.setTime(date);

                            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("d/MM/yyyy");
                            fecha = simpleDateFormat.format(date);
                            wsInforme.getCells().get("" + c + "" + (8 + list)).setValue(fecha);
                        } else {
                            f_error = true;
                            wsInforme.getCells().get("" + c + "" + (8 + list)).setValue("");
                        }
                    } else if (value == 3) {
                        int days = 0;
                        if (f_error != true) {
                            while (dunning.before(cierre) || dunning.equals(cierre)) {
                                if (dunning.get(Calendar.DAY_OF_WEEK) != Calendar.SUNDAY && dunning.get(Calendar.DAY_OF_WEEK) != Calendar.SATURDAY) {
                                    days++;
                                }
                                dunning.add(Calendar.DATE, 1);
                            }
                            days--;
                        }
                        wsInforme.getCells().get("" + c + "" + (8 + list)).setValue(days);
                    }
                    else if (value == 4) {
                        int total_suspensiones = Integer.parseInt(dataSuspensiones.get(value).get(list));
                        wsInforme.getCells().get("" + c + "" + (8 + list)).setValue(total_suspensiones);
                    } else if (value == 5) {
                        try {
                            int valor_total_cartera = Integer.parseInt(dataSuspensiones.get(value).get(list));
                            wsInforme.getCells().get("" + c + "" + (8 + list)).setValue(valor_total_cartera);
                            style.setNumber(5);
                        } catch (Exception e) {
                            wsInforme.getCells().get("" + c + "" + (8 + list)).setValue(0);
                            style.setNumber(5);
                        }
                    } else if (value == 6) {
                        int efectivas = Integer.parseInt(dataSuspensiones.get(value).get(list));
                        wsInforme.getCells().get("" + c + "" + (8 + list)).setValue(efectivas);
                    } else if (value == 7) {
                        int pagos = Integer.parseInt(dataSuspensiones.get(value).get(list));
                        wsInforme.getCells().get("" + c + "" + (8 + list)).setValue(pagos);
                    } else if (value == 8) {
                        int conserva_estado = Integer.parseInt(dataSuspensiones.get(value).get(list));
                        wsInforme.getCells().get("" + c + "" + (8 + list)).setValue(conserva_estado);
                    } else if (value == 9) {
                        int otras_anomalias = Integer.parseInt(dataSuspensiones.get(value).get(list));
                        wsInforme.getCells().get("" + c + "" + (8 + list)).setValue(otras_anomalias);
                    } else if (value == 10) {
                        try {
                            int valor_cartera_impresa = Integer.parseInt(dataSuspensiones.get(value).get(list));
                            wsInforme.getCells().get("" + c + "" + (8 + list)).setValue(valor_cartera_impresa);
                            style.setNumber(5);
                        } catch (Exception e) {
                            wsInforme.getCells().get("" + c + "" + (8 + list)).setValue(0);
                            style.setNumber(5);
                        }
                    } else if (value == 11) {
                        try {
                            int valor_cartera_efectiva = Integer.parseInt(dataSuspensiones.get(value).get(list));
                            wsInforme.getCells().get("" + c + "" + (8 + list)).setValue(valor_cartera_efectiva);
                            style.setNumber(5);
                        } catch (Exception e) {
                            wsInforme.getCells().get("" + c + "" + (8 + list)).setValue(0);
                            style.setNumber(5);
                        }
                    } else if (value == 12) {
                        wsInforme.getCells().get("" + c + "" + (8 + list)).setFormula("=M" + (8 + list) + "/N" + (8 + list));
                        style.setNumber(9);
                    } else {
                        try {
                            int valor_cartera_excluida = Integer.parseInt(dataSuspensiones.get(value).get(list));
                            wsInforme.getCells().get("" + c + "" + (8 + list)).setValue(valor_cartera_excluida);
                            style.setNumber(5);
                        } catch (Exception e) {
                            wsInforme.getCells().get("" + c + "" + (8 + list)).setValue(0);
                            style.setNumber(5);
                        }
                    }

                    wsInforme.getCells().get("" + c + "" + (8 + list)).setStyle(style);

                    if (value < 13) {
                        value++;
                    }

                    if (list < (dataSuspensiones.get(0).size()-1) && c == 'P') {
                        c = 'B';
                        value = 0;
                        list++;
                    }
                    style = new Style();
                }

                //TOTALIZADOR
                wsInforme.getCells().get("C" + (8 + (list+1))).setValue("TOTAL");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,192,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("C"+(8 + (list+1))+":E" + (8 + (list+1) +1));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((7+(list+1)),2,2,3);
                style = new Style();
                //PROMEDIO SUSPENSIONES
                wsInforme.getCells().get("F" + (8 + (list+1))).setFormula("=CONCATENATE(ROUND(AVERAGE(F8:F"+(7+(list+1)+"),1), \" DIAS\")"));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,192,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("F"+(8 + (list+1))+":F" + (8 + (list+1)+1));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((7+(list+1)),5,2,1);
                style = new Style();
                //SUMA -> TOTAL SUSPENSIONES
                wsInforme.getCells().get("G" + (8 + (list+1))).setFormula("=SUM(G8:G"+(7+(list+1)+")"));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,192,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("G"+(8 + (list+1))+":G" + (8 + (list+1)+1));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((7+(list+1)),6,2,1);
                style = new Style();
                //SUMA -> CARTERA TOTAL
                wsInforme.getCells().get("H" + (8 + (list+1))).setFormula("=SUM(H8:H"+(7+(list+1)+")"));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,192,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.setVerticalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("H"+(8 + (list+1))+":H" + (9 + (list+1)));
                range.applyStyle(style, flag);
                style.setNumber(5);
                wsInforme.getCells().get("H" + (8 + (list+1))).setStyle(style);
                wsInforme.getCells().merge((7+(list+1)),7,2,1);
                style = new Style();
                //SUMA -> EFECTIVAS
                wsInforme.getCells().get("I" + (8 + (list+1))).setFormula("=SUM(I8:I"+(7+(list+1)+")"));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,192,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                wsInforme.getCells().get("I" + (8 + (list+1))).setStyle(style);
                style = new Style();
                //PORCENTAJE -> EFECTIVAS
                wsInforme.getCells().get("I" + (9 + (list+1))).setFormula("=I"+(8 + (list+1))+"/G"+(8+(list+1)+""));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,192,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setNumber(9);
                wsInforme.getCells().get("I" + (9 + (list+1))).setStyle(style);
                style = new Style();
                //SUMA -> PAGOS
                wsInforme.getCells().get("J" + (8 + (list+1))).setFormula("=SUM(J8:J"+(7+(list+1)+")"));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,192,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                wsInforme.getCells().get("J" + (8 + (list+1))).setStyle(style);
                style = new Style();
                //PORCENTAJE -> PAGOS
                wsInforme.getCells().get("J" + (9 + (list+1))).setFormula("=J"+(8 + (list+1))+"/G"+(8+(list+1)+""));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,192,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setNumber(9);
                wsInforme.getCells().get("J" + (9 + (list+1))).setStyle(style);
                style = new Style();
                //SUMA -> CONSERVA ESTADO
                wsInforme.getCells().get("K" + (8 + (list+1))).setFormula("=SUM(K8:K"+(7+(list+1)+")"));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,192,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                wsInforme.getCells().get("K" + (8 + (list+1))).setStyle(style);
                style = new Style();
                //PORCENTAJE -> CONSERVA ESTADO
                wsInforme.getCells().get("K" + (9 + (list+1))).setFormula("=K"+(8 + (list+1))+"/G"+(8+(list+1)+""));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,192,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setNumber(9);
                wsInforme.getCells().get("K" + (9 + (list+1))).setStyle(style);
                style = new Style();
                //SUMA -> OTRAS ANOMALIAS
                wsInforme.getCells().get("L" + (8 + (list+1))).setFormula("=SUM(L8:L"+(7+(list+1)+")"));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,192,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                wsInforme.getCells().get("L" + (8 + (list+1))).setStyle(style);
                style = new Style();
                //PORCENTAJE -> OTRAS ANOMALIAS
                wsInforme.getCells().get("L" + (9 + (list+1))).setFormula("=L"+(8 + (list+1))+"/G"+(8+(list+1)+""));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,192,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setNumber(9);
                wsInforme.getCells().get("L" + (9 + (list+1))).setStyle(style);
                style = new Style();
                //SUMA -> CARTERA EFECTIVA
                wsInforme.getCells().get("M" + (8 + (list+1))).setFormula("=SUM(M8:M"+(7+(list+1)+")"));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,192,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.setVerticalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("M"+(8 + (list+1))+":M" + (9 + (list+1)));
                range.applyStyle(style, flag);
                style.setNumber(5);
                wsInforme.getCells().get("M" + (8 + (list+1))).setStyle(style);
                wsInforme.getCells().merge((7+(list+1)),12,2,1);
                style = new Style();
                //SUMA -> CARTERA ENVIADA A TERRENO
                wsInforme.getCells().get("N" + (8 + (list+1))).setFormula("=SUM(N8:N"+(7+(list+1)+")"));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,192,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.setVerticalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("N"+(8 + (list+1))+":N" + (9 + (list+1)));
                range.applyStyle(style, flag);
                style.setNumber(5);
                wsInforme.getCells().get("N" + (8 + (list+1))).setStyle(style);
                wsInforme.getCells().merge((7+(list+1)),13,2,1);
                style = new Style();
                //PORCENTAJE -> CARTERA SUSPENDIDA
                wsInforme.getCells().get("O" + (8 + (list+1))).setFormula("=M"+(8 + (list+1))+"/N"+(8+(list+1)+""));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,192,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.setVerticalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("O"+(8 + (list+1))+":O" + (9 + (list+1)));
                range.applyStyle(style, flag);
                style.setNumber(9);
                wsInforme.getCells().get("O" + (8 + (list+1))).setStyle(style);
                wsInforme.getCells().merge((7+(list+1)),14,2,1);
                style = new Style();
                //SUMA -> CARTERA EXCLUIDA
                wsInforme.getCells().get("P" + (8 + (list+1))).setFormula("=SUM(P8:P"+(7+(list+1)+")"));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,192,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.setVerticalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("P"+(8 + (list+1))+":P" + (9 + (list+1)));
                range.applyStyle(style, flag);
                style.setNumber(5);
                wsInforme.getCells().get("P" + (8 + (list+1))).setStyle(style);
                wsInforme.getCells().merge((7+(list+1)),15,2,1);
                style = new Style();
                //VALOR DE SUSPENSION
                wsInforme.getCells().get("C" + (10 + (list+1))).setValue("VALOR DE SUSPENSION");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getFont().setColor(Color.getWhite());
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("C"+(10 + (list+1))+":E" + (10 + (list+1)));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((9+(list+1)),2,1,3);
                style = new Style();
                //CELDA -> VALOR DE SUSPENSION
                wsInforme.getCells().get("F" + (10 + (list+1))).setValue(14000);
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,192,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("F"+(10 + (list+1))+":H" + (10 + (list+1)));
                range.applyStyle(style, flag);
                style.setNumber(5);
                wsInforme.getCells().get("F" + (10 + (list+1))).setStyle(style);
                wsInforme.getCells().merge((9+(list+1)),5,1,3);
                style = new Style();
                //VALOR TOTAL RECAUDADO
                wsInforme.getCells().get("C" + (11 + (list+1))).setValue("VALOR TOTAL RECAUDADO");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getFont().setColor(Color.getWhite());
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("C"+(11 + (list+1))+":E" + (11 + (list+1)));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((10+(list+1)),2,1,3);
                style = new Style();
                //CELDA -> VALOR RECAUDADO
                wsInforme.getCells().get("F" + (11 + (list+1))).setValue(0);
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,192,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("F"+(11 + (list+1))+":H" + (11 + (list+1)));
                range.applyStyle(style, flag);
                style.setNumber(5);
                wsInforme.getCells().get("F" + (11 + (list+1))).setStyle(style);
                wsInforme.getCells().merge((10+(list+1)),5,1,3);
                style = new Style();
                //VALOR TOTAL EJECUCION
                wsInforme.getCells().get("C" + (12 + (list+1))).setValue("VALOR TOTAL EJECUCION");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getFont().setColor(Color.getWhite());
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("C"+(12 + (list+1))+":E" + (12 + (list+1)));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((11+(list+1)),2,1,3);
                style = new Style();
                //CELDA -> VALOR EJECUCION
                wsInforme.getCells().get("F" + (12 + (list+1))).setFormula("=F" + (10 + (list+1)) + "*I" + (8 + (list+1)));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,192,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("F"+(12 + (list+1))+":H" + (12 + (list+1)));
                range.applyStyle(style, flag);
                style.setNumber(5);
                wsInforme.getCells().get("F" + (12 + (list+1))).setStyle(style);
                wsInforme.getCells().merge((11+(list+1)),5,1,3);
                style = new Style();
                //TOTAL RECAUDADO + EJECUTADO
                wsInforme.getCells().get("C" + (13 + (list+1))).setValue("TOTAL RECAUDADO + EJECUTADO");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getFont().setColor(Color.getWhite());
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("C"+(13 + (list+1))+":E" + (13 + (list+1)));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((12+(list+1)),2,1,3);
                style = new Style();
                //CELDA -> RECAUDADO + EJECUTADO
                wsInforme.getCells().get("F" + (13 + (list+1))).setFormula("=F" + (11 + (list+1)) + "+F" + (12 + (list+1)));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,192,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("F"+(13 + (list+1))+":H" + (13 + (list+1)));
                range.applyStyle(style, flag);
                style.setNumber(5);
                wsInforme.getCells().get("F" + (13 + (list+1))).setStyle(style);
                wsInforme.getCells().merge((12+(list+1)),5,1,3);
                style = new Style();
                //PORCENTAJE RECAUDADO
                wsInforme.getCells().get("C" + (14 + (list+1))).setValue("PORCENTAJE RECAUDADO");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getFont().setColor(Color.getWhite());
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("C"+(14 + (list+1))+":E" + (14 + (list+1)));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((13+(list+1)),2,1,3);
                style = new Style();
                //CELDA -> PORCENTAJE RECAUDADO
                wsInforme.getCells().get("F" + (14 + (list+1))).setFormula("=F" + (11 + (list+1)) + "/M" + (8 + (list+1)));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,192,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("F"+(14 + (list+1))+":H" + (14 + (list+1)));
                range.applyStyle(style, flag);
                style.setNumber(9);
                wsInforme.getCells().get("F" + (14 + (list+1))).setStyle(style);
                wsInforme.getCells().merge((13+(list+1)),5,1,3);
                style = new Style();
                //OBSERVACION
                wsInforme.getCells().get("I" + (10 + (list+1))).setValue("OBSERVACIÓN: ");
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("I"+(10 + (list+1))+":P" + (14 + (list+1)));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((9+(list+1)),8,5,8);
                style = new Style();

                //TABLA TAPONAMIENTOS
                wsInforme.getCells().get("C" + (15+(list+1))).setValue("TAPONAMIENTOS");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,255,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("C" + (15+(list+1))+ ":P" + (15+(list+1)));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((14+(list+1)),2,1,14);
                style = new Style();
                //PORCION
                wsInforme.getCells().get("C" + (16+(list+1))).setValue("PORCION");
                style.getFont().setBold(true);
                style.getFont().setColor(Color.getWhite());
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("C" + (16+(list+1))+ ":C" + (17+(list+1)));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((15+(list+1)),2,2,1);
                style = new Style();
                //DUNNING
                wsInforme.getCells().get("D" + (16+(list+1))).setValue("FECHA DUNNING");
                style.getFont().setBold(true);
                style.getFont().setColor(Color.getWhite());
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("D" + (16+(list+1))+ ":D" + (17+(list+1)));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((15+(list+1)),3,2,1);
                style = new Style();
                //CIERRE
                wsInforme.getCells().get("E" + (16+(list+1))).setValue("FECHA CIERRE");
                style.getFont().setBold(true);
                style.getFont().setColor(Color.getWhite());
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("E" + (16+(list+1))+ ":E" + (17+(list+1)));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((15+(list+1)),4,2,1);
                style = new Style();
                //PROMEDIO
                wsInforme.getCells().get("F" + (16+(list+1))).setValue("PROMEDIO\n(días)");
                style.getFont().setBold(true);
                style.getFont().setColor(Color.getWhite());
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setHorizontalAlignment(TextAlignmentType.DISTRIBUTED);
                range = wsInforme.getCells().createRange("F" + (16+(list+1))+ ":F" + (17+(list+1)));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((15+(list+1)),5,2,1);
                style = new Style();
                //TOTAL SUSPENSIONES
                wsInforme.getCells().get("G" + (16+(list+1))).setValue("TOTAL\nTAPONAMIENTOS");
                style.getFont().setBold(true);
                style.getFont().setColor(Color.getWhite());
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setHorizontalAlignment(TextAlignmentType.DISTRIBUTED);
                range = wsInforme.getCells().createRange("G" + (16+(list+1))+ ":G" + (17+(list+1)));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((15+(list+1)),6,2,1);
                style = new Style();
                //VALOR TOTAL CARTERA
                wsInforme.getCells().get("H" + (16+(list+1))).setValue("CARTERA TOTAL");
                style.getFont().setBold(true);
                style.getFont().setColor(Color.getWhite());
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                range = wsInforme.getCells().createRange("H" + (16+(list+1))+ ":H" + (17+(list+1)));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((15+(list+1)),7,2,1);
                style = new Style();
                //RESULTADO
                wsInforme.getCells().get("I" + (16+(list+1))).setValue("RESULTADO");
                style.getFont().setBold(true);
                style.getFont().setColor(Color.getWhite());
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("I" + (16+(list+1))+ ":L" + (16+(list+1)));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((15+(list+1)),8,1,4);
                style = new Style();
                //R -> EFECTIVAS
                wsInforme.getCells().get("I" + (17+(list+1))).setValue("EFECTIVAS");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,230,153));
                style.setPattern(BackgroundType.SOLID);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.setVerticalAlignment(TextAlignmentType.CENTER);
                wsInforme.getCells().get("I" + (17+(list+1))).setStyle(style);
                style = new Style();
                //R -> PAGOS
                wsInforme.getCells().get("J" + (17+(list+1))).setValue("PAGOS");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,230,153));
                style.setPattern(BackgroundType.SOLID);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.setVerticalAlignment(TextAlignmentType.CENTER);
                wsInforme.getCells().get("J" + (17+(list+1))).setStyle(style);
                style = new Style();
                //R -> CONSERVA ESTADO
                wsInforme.getCells().get("K" + (17+(list+1))).setValue("CONSERVA ESTADO");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,230,153));
                style.setPattern(BackgroundType.SOLID);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setHorizontalAlignment(TextAlignmentType.DISTRIBUTED);
                style.setVerticalAlignment(TextAlignmentType.CENTER);
                wsInforme.getCells().get("K" + (17+(list+1))).setStyle(style);
                style = new Style();
                //R -> OTRAS ANOMALIAS
                wsInforme.getCells().get("L" + (17+(list+1))).setValue("OTRAS ANOMALIAS");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,230,153));
                style.setPattern(BackgroundType.SOLID);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setHorizontalAlignment(TextAlignmentType.DISTRIBUTED);
                style.setVerticalAlignment(TextAlignmentType.CENTER);
                wsInforme.getCells().get("L" + (17+(list+1))).setStyle(style);
                style = new Style();
                //VALOR CARTERA EFECTIVA
                wsInforme.getCells().get("M" + (16+(list+1))).setValue("VALOR\nCARTERA\nEFECTIVA");
                style.getFont().setBold(true);
                style.getFont().setColor(Color.getWhite());
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setHorizontalAlignment(TextAlignmentType.DISTRIBUTED);
                style.setVerticalAlignment(TextAlignmentType.CENTER);
                range = wsInforme.getCells().createRange("M" + (16+(list+1)) + ":M" + (17+(list+1)));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((15+(list+1)),12,2,1);
                style = new Style();
                //VALOR CARTERA ENVIADA A TERRENO
                wsInforme.getCells().get("N" + (16+(list+1))).setValue("CARTERA\nENVIADA\nTERRENO");
                style.getFont().setBold(true);
                style.getFont().setColor(Color.getWhite());
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setHorizontalAlignment(TextAlignmentType.DISTRIBUTED);
                style.setVerticalAlignment(TextAlignmentType.CENTER);
                range = wsInforme.getCells().createRange("N" + (16+(list+1)) + ":N" + (17+(list+1)));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((15+(list+1)),13,2,1);
                style = new Style();
                //PORCENTAJE CARTERA SUSPENDIDA
                wsInforme.getCells().get("O" + (16+(list+1))).setValue("PORCENTAJE\nCARTERA\nTAPONADA");
                style.getFont().setBold(true);
                style.getFont().setColor(Color.getWhite());
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setHorizontalAlignment(TextAlignmentType.DISTRIBUTED);
                style.setVerticalAlignment(TextAlignmentType.CENTER);
                range = wsInforme.getCells().createRange("O" + (16+(list+1)) + ":O" + (17+(list+1)));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((15+(list+1)),14,2,1);
                style = new Style();
                //VALOR CARTERA EXCLUIDA
                wsInforme.getCells().get("P" + (16+(list+1))).setValue("CARTERA\nEXCLUIDA");
                style.getFont().setBold(true);
                style.getFont().setColor(Color.getWhite());
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setHorizontalAlignment(TextAlignmentType.DISTRIBUTED);
                style.setVerticalAlignment(TextAlignmentType.CENTER);
                range = wsInforme.getCells().createRange("P" + (16+(list+1)) + ":P" + (17+(list+1)));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((15+(list+1)),15,2,1);
                style = new Style();

                //DATA
                int list2 = 17 + list+1;
                list = 0;
                value = 0;

                dunning = Calendar.getInstance();
                cierre = Calendar.getInstance();
                f_error = false;

                for (c = 'C'; c <= 'P'; c++) {
                    style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                    style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                    style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                    style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                    style.setHorizontalAlignment(TextAlignmentType.CENTER);

                    if (value == 0) {
                        wsInforme.getCells().get("" + c + "" + (list2+(list+1))).setValue(dataTaponamientos.get(value).get(list));
                    } else if (value == 1) {
                        String fecha = dataTaponamientos.get(1).get(list);
                        if (fecha != null) {
                            f_error = false;
                            DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
                            Date date = dateFormat.parse(fecha);

                            dunning.setTime(date);

                            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("d/MM/yyyy");
                            fecha = simpleDateFormat.format(date);
                            wsInforme.getCells().get("" + c + "" + (list2 + (list + 1))).setValue(fecha);

                        } else {
                            f_error = true;
                            wsInforme.getCells().get("" + c + "" + (list2 + (list + 1))).setValue("");
                        }
                    } else if (value == 2) {
                        String fecha = dataTaponamientos.get(2).get(list);
                        if (fecha != null) {
                            f_error = false;
                            DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
                            Date date = dateFormat.parse(fecha);

                            cierre.setTime(date);

                            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("d/MM/yyyy");
                            fecha = simpleDateFormat.format(date);
                            wsInforme.getCells().get("" + c + "" + (list2 + (list + 1))).setValue(fecha);
                        } else {
                            f_error = true;
                            wsInforme.getCells().get("" + c + "" + (list2 + (list + 1))).setValue("");
                        }
                    } else if (value == 3) {
                        int days = 0;
                        if (f_error != true) {
                        while (dunning.before(cierre) || dunning.equals(cierre)) {
                            if (dunning.get(Calendar.DAY_OF_WEEK) != Calendar.SUNDAY && dunning.get(Calendar.DAY_OF_WEEK) != Calendar.SATURDAY) {
                                days++;
                            }
                            dunning.add(Calendar.DATE, 1);
                        }
                        days--;
                        }
                        wsInforme.getCells().get("" + c + "" + (list2+(list+1))).setValue(days);
                    }
                    else if (value == 4) {
                        int total_suspensiones = Integer.parseInt(dataTaponamientos.get(value).get(list));
                        wsInforme.getCells().get("" + c + "" + (list2+(list+1))).setValue(total_suspensiones);
                    } else if (value == 5) {
                        try {
                            int valor_total_cartera = Integer.parseInt(dataTaponamientos.get(value).get(list));
                            wsInforme.getCells().get("" + c + "" + (list2 + (list + 1))).setValue(valor_total_cartera);
                            style.setNumber(5);
                        } catch (Exception e) {
                            wsInforme.getCells().get("" + c + "" + (list2 + (list + 1))).setValue(0);
                            style.setNumber(5);
                        }
                    } else if (value == 6) {
                        int efectivas = Integer.parseInt(dataTaponamientos.get(value).get(list));
                        wsInforme.getCells().get("" + c + "" + (list2+(list+1))).setValue(efectivas);
                    } else if (value == 7) {
                        int pagos = Integer.parseInt(dataTaponamientos.get(value).get(list));
                        wsInforme.getCells().get("" + c + "" + (list2+(list+1))).setValue(pagos);
                    } else if (value == 8) {
                        int conserva_estado = Integer.parseInt(dataTaponamientos.get(value).get(list));
                        wsInforme.getCells().get("" + c + "" + (list2+(list+1))).setValue(conserva_estado);
                    } else if (value == 9) {
                        int otras_anomalias = Integer.parseInt(dataTaponamientos.get(value).get(list));
                        wsInforme.getCells().get("" + c + "" + (list2+(list+1))).setValue(otras_anomalias);
                    } else if (value == 10) {
                        try {
                            int valor_cartera_impresa = Integer.parseInt(dataTaponamientos.get(value).get(list));
                            wsInforme.getCells().get("" + c + "" + (list2 + (list + 1))).setValue(valor_cartera_impresa);
                            style.setNumber(5);
                        } catch (Exception e) {
                            wsInforme.getCells().get("" + c + "" + (list2 + (list + 1))).setValue(0);
                            style.setNumber(5);
                        }
                    } else if (value == 11) {
                        try {
                            int valor_cartera_efectiva = Integer.parseInt(dataTaponamientos.get(value).get(list));
                            wsInforme.getCells().get("" + c + "" + (list2 + (list + 1))).setValue(valor_cartera_efectiva);
                            style.setNumber(5);
                        } catch (Exception e) {
                            wsInforme.getCells().get("" + c + "" + (list2 + (list + 1))).setValue(0);
                            style.setNumber(5);
                        }
                    } else if (value == 12) {
                        wsInforme.getCells().get("" + c + "" + (list2+(list+1))).setFormula("=M" + (list2+(list+1)) + "/N" + (list2+(list+1)));
                        style.setNumber(9);
                    } else {
                        try {
                            int valor_cartera_excluida = Integer.parseInt(dataTaponamientos.get(value).get(list));
                            wsInforme.getCells().get("" + c + "" + (list2 + (list + 1))).setValue(valor_cartera_excluida);
                            style.setNumber(5);
                        } catch (Exception e) {
                            wsInforme.getCells().get("" + c + "" + (list2 + (list + 1))).setValue(0);
                            style.setNumber(5);
                        }
                    }

                    wsInforme.getCells().get("" + c + "" + (list2+(list+1))).setStyle(style);

                    if (value < 13) {
                        value++;
                    }

                    if (list < (dataTaponamientos.get(0).size()-1) && c == 'P') {
                        c = 'B';
                        value = 0;
                        list++;
                    }
                    style = new Style();
                }

                //TOTALIZADOR
                wsInforme.getCells().get("C" + (list2+(list+1)+1)).setValue("TOTAL");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,255,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("C"+(list2+(list+1)+1)+":E" + (list2+(list+1)+2));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((list2+(list+1)),2,2,3);
                style = new Style();
                //PROMEDIO SUSPENSIONES
                wsInforme.getCells().get("F" + (list2+(list+1)+1)).setFormula("=CONCATENATE(ROUND(AVERAGE(F"+ (list2+1) + ":F" + (list2+(list+1)) + "),1), \" DIAS\")");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,255,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("F"+(list2+(list+1)+1)+":F" + (list2+(list+1)+2));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((list2+(list+1)),5,2,1);
                style = new Style();
                //SUMA -> TOTAL TAPONAMIENTOS
                wsInforme.getCells().get("G" + (list2+(list+1)+1)).setFormula("=SUM(G"+(list2+1)+":G"+(list2+(list+1)+")"));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,255,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("G"+(list2+(list+1)+1)+":G" + (list2+(list+1)+2));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((list2+(list+1)),6,2,1);
                style = new Style();
                //SUMA -> CARTERA TOTAL
                wsInforme.getCells().get("H" + (list2+(list+1)+1)).setFormula("=SUM(H"+(list2+1)+":H"+(list2+(list+1)+")"));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,255,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.setVerticalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("H"+(list2+(list+1)+1)+":H" + (list2+(list+1)+2));
                range.applyStyle(style, flag);
                style.setNumber(5);
                wsInforme.getCells().get("H" + (list2+(list+1)+1)).setStyle(style);
                wsInforme.getCells().merge((list2+(list+1)),7,2,1);
                style = new Style();
                //SUMA -> EFECTIVAS
                wsInforme.getCells().get("I" + (list2+(list+1)+1)).setFormula("=SUM(I"+(list2+1)+":I"+(list2+(list+1)+")"));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,255,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                wsInforme.getCells().get("I" + (list2+(list+1)+1)).setStyle(style);
                style = new Style();
                //PORCENTAJE -> EFECTIVAS
                wsInforme.getCells().get("I" + (list2+(list+1)+2)).setFormula("=I"+(list2+(list+1)+1)+"/G"+(list2+(list+1)+1)+"");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,255,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setNumber(9);
                wsInforme.getCells().get("I" + (list2+(list+1)+2)).setStyle(style);
                style = new Style();
                //SUMA -> PAGOS
                wsInforme.getCells().get("J" + (list2+(list+1)+1)).setFormula("=SUM(J"+(list2+1)+":J"+(list2+(list+1)+")"));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,255,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                wsInforme.getCells().get("J" + (list2+(list+1)+1)).setStyle(style);
                style = new Style();
                //PORCENTAJE -> PAGOS
                wsInforme.getCells().get("J" + (list2+(list+1)+2)).setFormula("=J"+(list2+(list+1)+1)+"/G"+(list2+(list+1)+1)+"");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,255,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setNumber(9);
                wsInforme.getCells().get("J" + (list2+(list+1)+2)).setStyle(style);
                style = new Style();
                //SUMA -> CONSERVA ESTADO
                wsInforme.getCells().get("K" + (list2+(list+1)+1)).setFormula("=SUM(K"+(list2+1)+":K"+(list2+(list+1)+")"));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,255,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                wsInforme.getCells().get("K" + (list2+(list+1)+1)).setStyle(style);
                style = new Style();
                //PORCENTAJE -> CONSERVA ESTADO
                wsInforme.getCells().get("K" + (list2+(list+1)+2)).setFormula("=K"+(list2+(list+1)+1)+"/G"+(list2+(list+1)+1));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,255,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setNumber(9);
                wsInforme.getCells().get("K" + (list2+(list+1)+2)).setStyle(style);
                style = new Style();
                //SUMA -> OTRAS ANOMALIAS
                wsInforme.getCells().get("L" + (list2+(list+1)+1)).setFormula("=SUM(L"+(list2+1)+":L"+(list2+(list+1)+")"));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,255,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                wsInforme.getCells().get("L" + (list2+(list+1)+1)).setStyle(style);
                style = new Style();
                //PORCENTAJE -> OTRAS ANOMALIAS
                wsInforme.getCells().get("L" + (list2+(list+1)+2)).setFormula("=L"+(list2+(list+1)+1)+"/G"+(list2+(list+1)+1));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,255,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setNumber(9);
                wsInforme.getCells().get("L" + (list2+(list+1)+2)).setStyle(style);
                style = new Style();
                //SUMA -> CARTERA EFECTIVA
                wsInforme.getCells().get("M" + (list2+(list+1)+1)).setFormula("=SUM(M"+(list2+1)+":M"+(list2+(list+1)+")"));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,255,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.setVerticalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("M"+(list2+(list+1)+1)+":M" + (list2+(list+1)+2));
                range.applyStyle(style, flag);
                style.setNumber(5);
                wsInforme.getCells().get("M" + (list2+(list+1)+1)).setStyle(style);
                wsInforme.getCells().merge((list2+(list+1)),12,2,1);
                style = new Style();
                //SUMA -> CARTERA ENVIADA A TERRENO
                wsInforme.getCells().get("N" + (list2+(list+1)+1)).setFormula("=SUM(N"+(list2+1)+":N"+(list2+(list+1)+")"));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,255,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.setVerticalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("N"+(list2+(list+1)+1)+":N" + (list2+(list+1)+2));
                range.applyStyle(style, flag);
                style.setNumber(5);
                wsInforme.getCells().get("N" + (list2+(list+1)+1)).setStyle(style);
                wsInforme.getCells().merge((list2+(list+1)),13,2,1);
                style = new Style();
                //PORCENTAJE -> CARTERA SUSPENDIDA
                wsInforme.getCells().get("O" + (list2+(list+1)+1)).setFormula("=M"+(list2+(list+1)+1)+"/N"+(list2+(list+1)+1));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,255,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.setVerticalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("O"+(list2+(list+1)+1)+":O" + (list2+(list+1)+2));
                range.applyStyle(style, flag);
                style.setNumber(9);
                wsInforme.getCells().get("O" + (list2+(list+1)+1)).setStyle(style);
                wsInforme.getCells().merge((list2+(list+1)),14,2,1);
                style = new Style();
                //SUMA -> CARTERA EXCLUIDA
                wsInforme.getCells().get("P" + (list2+(list+1)+1)).setFormula("=SUM(P"+(list2+1)+":P"+(list2+(list+1)+")"));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,255,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.setVerticalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("P"+(list2+(list+1)+1)+":P" + (list2+(list+1)+2));
                range.applyStyle(style, flag);
                style.setNumber(5);
                wsInforme.getCells().get("P" + (list2+(list+1)+1)).setStyle(style);
                wsInforme.getCells().merge((list2+(list+1)),15,2,1);
                style = new Style();
                //VALOR DE SUSPENSION
                wsInforme.getCells().get("C" + (list2+(list+1)+3)).setValue("VALOR DE TAPONAMIENTO");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getFont().setColor(Color.getWhite());
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("C"+(list2+(list+1)+3)+":E" + (list2+(list+1)+3));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((list2+(list+1)+2),2,1,3);
                style = new Style();
                //CELDA -> VALOR DE TAPONAMIENTO
                wsInforme.getCells().get("F" + (list2+(list+1)+3)).setValue(24000);
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,255,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("F"+(list2+(list+1)+3)+":H" + (list2+(list+1)+3));
                range.applyStyle(style, flag);
                style.setNumber(5);
                wsInforme.getCells().get("F" + (list2+(list+1)+3)).setStyle(style);
                wsInforme.getCells().merge((list2+(list+1)+2),5,1,3);
                style = new Style();
                //VALOR TOTAL RECAUDADO
                wsInforme.getCells().get("C" + (list2+(list+1)+4)).setValue("VALOR TOTAL RECAUDADO");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getFont().setColor(Color.getWhite());
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("C"+(list2+(list+1)+4)+":E" + (list2+(list+1)+4));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((list2+(list+1)+3),2,1,3);
                style = new Style();
                //CELDA -> VALOR RECAUDADO
                wsInforme.getCells().get("F" + (list2+(list+1)+4)).setValue(0);
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,255,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("F"+(list2+(list+1)+4)+":H" + (list2+(list+1)+4));
                range.applyStyle(style, flag);
                style.setNumber(5);
                wsInforme.getCells().get("F" + (list2+(list+1)+4)).setStyle(style);
                wsInforme.getCells().merge((list2+(list+1)+3),5,1,3);
                style = new Style();
                //VALOR TOTAL EJECUCION
                wsInforme.getCells().get("C" + (list2+(list+1)+5)).setValue("VALOR TOTAL EJECUCION");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getFont().setColor(Color.getWhite());
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("C"+(list2+(list+1)+5)+":E" + (list2+(list+1)+5));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((list2+(list+1)+4),2,1,3);
                style = new Style();
                //CELDA -> VALOR EJECUCION
                wsInforme.getCells().get("F" + (list2+(list+1)+5)).setFormula("=F" + (list2+(list+1)+3) + "*I" + (list2+(list+1)+1));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,255,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("F"+(list2+(list+1)+5)+":H" + (list2+(list+1)+5));
                range.applyStyle(style, flag);
                style.setNumber(5);
                wsInforme.getCells().get("F" + (list2+(list+1)+5)).setStyle(style);
                wsInforme.getCells().merge((list2+(list+1)+4),5,1,3);
                style = new Style();
                //TOTAL RECAUDADO + EJECUTADO
                wsInforme.getCells().get("C" + (list2+(list+1)+6)).setValue("TOTAL RECAUDADO + EJECUTADO");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getFont().setColor(Color.getWhite());
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("C"+(list2+(list+1)+6)+":E" + (list2+(list+1)+6));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((list2+(list+1)+5),2,1,3);
                style = new Style();
                //CELDA -> RECAUDADO + EJECUTADO
                wsInforme.getCells().get("F" + (list2+(list+1)+6)).setFormula("=F" + (list2+(list+1)+4) + "+F" + (list2+(list+1)+5));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,255,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("F"+(list2+(list+1)+6)+":H" + (list2+(list+1)+6));
                range.applyStyle(style, flag);
                style.setNumber(5);
                wsInforme.getCells().get("F" + (list2+(list+1)+6)).setStyle(style);
                wsInforme.getCells().merge((list2+(list+1)+5),5,1,3);
                style = new Style();
                //PORCENTAJE RECAUDADO
                wsInforme.getCells().get("C" + (list2+(list+1)+7)).setValue("PORCENTAJE RECAUDADO");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getFont().setColor(Color.getWhite());
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("C"+(list2+(list+1)+7)+":E" + (list2+(list+1)+7));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((list2+(list+1)+6),2,1,3);
                style = new Style();
                //CELDA -> PORCENTAJE RECAUDADO
                wsInforme.getCells().get("F" + (list2+(list+1)+7)).setFormula("=F" + (list2+(list+1)+4) + "/M" + (list2+(list+1)+1));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,255,0));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("F"+(list2+(list+1)+7)+":H" + (list2+(list+1)+7));
                range.applyStyle(style, flag);
                style.setNumber(9);
                wsInforme.getCells().get("F" + (list2+(list+1)+7)).setStyle(style);
                wsInforme.getCells().merge((list2+(list+1)+6),5,1,3);
                style = new Style();
                //OBSERVACION
                wsInforme.getCells().get("I" + (list2+(list+1)+3)).setValue("OBSERVACIÓN: ");
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("I"+(list2+(list+1)+3)+":P" + (list2+(list+1)+7));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((list2+(list+1)+2),8,5,8);
                style = new Style();

                int list3 = list + list2;
                //TABLA REINSTALACIONES
                wsInforme.getCells().get("C" + (list3 + 9)).setValue("REINSTALACIONES");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(146,208,80));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("C" + (list3 + 9)+ ":P" + (list3 + 9));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((list3 + 8),2,1,14);
                style = new Style();
                //FECHA REINSTALACION
                wsInforme.getCells().get("C" + (list3 + 10)).setValue("FECHA REINSTALACION");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getFont().setColor(Color.getWhite());
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("C" + (list3 + 10)+ ":D" + (list3 + 11));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((list3 + 9),2,2,2);
                style = new Style();
                //FECHA CIERRE
                wsInforme.getCells().get("E" + (list3 + 10)).setValue("FECHA CIERRE");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getFont().setColor(Color.getWhite());
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("E" + (list3 + 10)+ ":F" + (list3 + 11));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((list3 + 9),4    ,2,2);
                style = new Style();
                //PROMEDIO DIAS
                wsInforme.getCells().get("G" + (list3 + 10)).setValue("PROMEDIO (DIAS)");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getFont().setColor(Color.getWhite());
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("G" + (list3 + 10)+ ":G" + (list3 + 11));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((list3 + 9),6    ,2,1);
                style = new Style();
                //TOTAL REINSTALACIONES
                wsInforme.getCells().get("H" + (list3 + 10)).setValue("TOTAL\nREINSTALACIONES");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getFont().setColor(Color.getWhite());
                style.setHorizontalAlignment(TextAlignmentType.DISTRIBUTED);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("H" + (list3 + 10)+ ":H" + (list3 + 11));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((list3 + 9),7    ,2,1);
                style = new Style();
                //RESULTADO
                wsInforme.getCells().get("I" + (list3 + 10)).setValue("RESULTADO");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getFont().setColor(Color.getWhite());
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("I" + (list3 + 10)+ ":L" + (list3 + 10));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((list3 + 9),8    ,1,4);
                style = new Style();
                //R -> EFECTIVAS
                wsInforme.getCells().get("I" + (list3 + 11)).setValue("EFECTIVAS");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,230,153));
                style.setPattern(BackgroundType.SOLID);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.setVerticalAlignment(TextAlignmentType.CENTER);
                range = wsInforme.getCells().createRange("I" + (list3 + 11)+ ":J" + (list3 + 11));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((list3 + 10),8,1,2);
                style = new Style();
                //R -> INEFECTIVAS
                wsInforme.getCells().get("K" + (list3 + 11)).setValue("INEFECTIVAS");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(255,230,153));
                style.setPattern(BackgroundType.SOLID);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.setVerticalAlignment(TextAlignmentType.CENTER);
                range = wsInforme.getCells().createRange("K" + (list3 + 11)+ ":L" + (list3 + 11));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((list3 + 10),10,1,2);
                style = new Style();
                //PORCION
                wsInforme.getCells().get("O" + (list3 + 10)).setValue("PORCION");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getFont().setColor(Color.getWhite());
                style.setHorizontalAlignment(TextAlignmentType.DISTRIBUTED);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("O" + (list3 + 10)+ ":O" + (list3 + 11));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((list3 + 9),14,2,1);
                style = new Style();
                //TOTAL REINSTALACIONES
                wsInforme.getCells().get("P" + (list3 + 10)).setValue("TOTAL\nREINSTALACIONES");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getFont().setColor(Color.getWhite());
                style.setHorizontalAlignment(TextAlignmentType.DISTRIBUTED);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("P" + (list3 + 10)+ ":P" + (list3 + 11));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((list3 + 9),15,2,1);
                style = new Style();
                //DATA
                list = 0;
                value = 0;

                for (c = 'C'; c <= 'L'; c++) {
                    style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                    style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                    style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                    style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                    style.setHorizontalAlignment(TextAlignmentType.CENTER);

                    if (value == 0) {
                        String fecha = dataReinstalaciones.get(0).get(list);

                        DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
                        Date date = dateFormat.parse(fecha);

                        dunning.setTime(date);
                        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("d/MM/yyyy");
                        fecha = simpleDateFormat.format(date);

                        wsInforme.getCells().get("" + c + "" + ((list3 + 12)+list)).setValue(fecha);
                        range = wsInforme.getCells().createRange("" + c + "" + ((list3 + 12)+list) + ":" + (c++) + "" + ((list3 + 12)+list));
                        range.applyStyle(style, flag);
                        wsInforme.getCells().merge(((list3 + 11)+list),2,1,2);
                    } else if (value == 1) {
                        String fecha = dataReinstalaciones.get(1).get(list);

                        DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
                        Date date = dateFormat.parse(fecha);

                        cierre.setTime(date);
                        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("d/MM/yyyy");
                        fecha = simpleDateFormat.format(date);

                        wsInforme.getCells().get("" + c + "" + ((list3 + 12)+list)).setValue(fecha);
                        range = wsInforme.getCells().createRange("" + c + "" + ((list3 + 12)+list) + ":" + (c++) + ((list3 + 12)+list));
                        range.applyStyle(style, flag);
                        wsInforme.getCells().merge(((list3 + 11)+list),4,1,2);
                    } else if (value == 2) {
                        int days = 0;

                        while (dunning.before(cierre) || dunning.equals(cierre)) {
                            if (dunning.get(Calendar.DAY_OF_WEEK) != Calendar.SUNDAY && dunning.get(Calendar.DAY_OF_WEEK) != Calendar.SATURDAY) {
                                days++;
                            }
                            dunning.add(Calendar.DATE, 1);
                        }
                        days--;

                        wsInforme.getCells().get("" + c + "" + ((list3 + 12)+list)).setValue(days);
                    } else if (value == 3) {
                        wsInforme.getCells().get("" + c + "" + ((list3 + 12)+list)).setValue(Integer.parseInt(dataReinstalaciones.get(value).get(list)));
                    }
                    else if (value == 4) {
                        wsInforme.getCells().get("" + c + "" + ((list3 + 12)+list)).setValue(Integer.parseInt(dataReinstalaciones.get(value).get(list)));
                        range = wsInforme.getCells().createRange("" + c + "" + ((list3 + 12)+list) + ":" + (c++) + ((list3 + 12)+list));
                        range.applyStyle(style, flag);
                        wsInforme.getCells().merge(((list3 + 11)+list),8,1,2);
                    } else {
                        wsInforme.getCells().get("" + c + "" + ((list3 + 12)+list)).setValue(Integer.parseInt(dataReinstalaciones.get(value).get(list)));
                        range = wsInforme.getCells().createRange("" + c + "" + ((list3 + 12)+list) + ":" + (c++) + ((list3 + 12)+list));
                        range.applyStyle(style, flag);
                        wsInforme.getCells().merge(((list3 + 11)+list),10,1,2);
                    }

                    if (c != 'C' && c != 'E' && c != 'I' && c != 'K') {
                        wsInforme.getCells().get("" + c + "" + ((list3 + 12) + list)).setStyle(style);
                    }
                    if (value < 5) {
                        value++;
                    }

                    if (list < (dataReinstalaciones.get(0).size()-1) && c == 'L') {
                        c = 'B';
                        value = 0;
                        list++;
                    }
                    style = new Style();
                }

                //TOTALIZADOR
                wsInforme.getCells().get("C" + ((list3 + 13) + list)).setValue("TOTAL");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(146,208,80));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("C"+((list3 + 13) + list)+":F" + ((list3 + 14) + list));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge(((list3 + 12) + list),2,2,4);
                style = new Style();
                //PROMEDIO REINSTALACIONES
                wsInforme.getCells().get("G" + ((list3 + 13) + list)).setFormula("=CONCATENATE(ROUND(AVERAGE(G"+ ((list3 + 12)) + ":G" + ((list3 + 12) + list) + "),1), \" DIAS\")");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(146,208,80));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("G"+((list3 + 13) + list)+":G" + ((list3 + 14) + list));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge(((list3 + 12) + list),6,2,1);
                style = new Style();
                //SUMA -> TOTAL REINSTALACIONES
                wsInforme.getCells().get("H" + ((list3 + 13) + list)).setFormula("=SUM(H"+((list3 + 12))+":H"+((list3 + 12) + list)+")");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(146,208,80));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("H"+((list3 + 13) + list)+":H" + ((list3 + 14) + list));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge(((list3 + 12) + list),7,2,1);
                style = new Style();
                //SUMA -> EFECTIVAS
                wsInforme.getCells().get("I" + ((list3 + 13) + list)).setFormula("=SUM(I"+((list3 + 12))+":I"+((list3 + 12) + list)+")");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(146,208,80));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("I"+((list3 + 13) + list)+":J" + ((list3 + 13) + list));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge(((list3 + 12) + list),8,1,2);
                style = new Style();
                //PORCENTAJE -> EFECTIVAS
                wsInforme.getCells().get("I" + ((list3 + 14) + list)).setFormula("=I"+((list3 + 13) + list)+"/H"+((list3 + 13) + list));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(146,208,80));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("I"+((list3 + 14) + list)+":J" + ((list3 + 14) + list));
                range.applyStyle(style, flag);
                style.setNumber(9);
                wsInforme.getCells().get("I" + ((list3 + 14) + list)).setStyle(style);
                wsInforme.getCells().merge(((list3 + 13) + list),8,1,2);
                style = new Style();
                //SUMA -> INEFECTIVAS
                wsInforme.getCells().get("K" + ((list3 + 13) + list)).setFormula("=SUM(K"+((list3 + 12))+":K"+((list3 + 12) + list)+")");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(146,208,80));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("K"+((list3 + 13) + list)+":L" + ((list3 + 13) + list));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge(((list3 + 12) + list),10,1,2);
                style = new Style();
                //PORCENTAJE -> INEFECTIVAS
                wsInforme.getCells().get("K" + ((list3 + 14) + list)).setFormula("=K"+((list3 + 13) + list)+"/H"+((list3 + 13) + list));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(146,208,80));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("K"+((list3 + 14) + list)+":L" + ((list3 + 14) + list));
                range.applyStyle(style, flag);
                style.setNumber(9);
                wsInforme.getCells().get("K" + ((list3 + 14) + list)).setStyle(style);
                wsInforme.getCells().merge(((list3 + 13) + list),10,1,2);
                style = new Style();
                //VALOR DE REINSTALACION
                wsInforme.getCells().get("C" + ((list3 + 15) + list)).setValue("VALOR DE REINSTALACION");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getFont().setColor(Color.getWhite());
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("C"+((list3 + 15) + list)+":F" + ((list3 + 15) + list));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge(((list3 + 14) + list),2,1,4);
                style = new Style();
                //CELDA -> VALOR DE REINSTALACION
                wsInforme.getCells().get("G" + ((list3 + 15) + list)).setValue(12000);
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(146,208,80));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("G"+((list3 + 15) + list)+":H" + ((list3 + 15) + list));
                range.applyStyle(style, flag);
                style.setNumber(5);
                wsInforme.getCells().get("G" + ((list3 + 15) + list)).setStyle(style);
                wsInforme.getCells().merge(((list3 + 14) + list),6,1,2);
                style = new Style();
                //VALOR TOTAL RECAUDADO
                wsInforme.getCells().get("C" + ((list3 + 16) + list)).setValue("VALOR TOTAL RECAUDADO");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getFont().setColor(Color.getWhite());
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("C"+((list3 + 16) + list)+":F" + ((list3 + 16) + list));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge(((list3 + 15) + list),2,1,4);
                style = new Style();
                //CELDA -> VALOR RECAUDADO
                wsInforme.getCells().get("G" + ((list3 + 16) + list)).setValue(0);
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(146,208,80));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("G"+((list3 + 16) + list)+":H" + ((list3 + 16) + list));
                range.applyStyle(style, flag);
                style.setNumber(5);
                wsInforme.getCells().get("G" + ((list3 + 16) + list)).setStyle(style);
                wsInforme.getCells().merge(((list3 + 15) + list),6,1,2);
                style = new Style();
                //VALOR TOTAL EJECUCION
                wsInforme.getCells().get("C" + ((list3 + 17) + list)).setValue("VALOR TOTAL EJECUCION");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getFont().setColor(Color.getWhite());
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("C"+((list3 + 17) + list)+":F" + ((list3 + 17) + list));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge(((list3 + 16) + list),2,1,4);
                style = new Style();
                //CELDA -> VALOR EJECUCION
                wsInforme.getCells().get("G" + ((list3 + 17) + list)).setFormula("=G" + ((list3 + 15) + list) + "*I" + ((list3 + 13) + list));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(146,208,80));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("G"+((list3 + 17) + list)+":H" + ((list3 + 17) + list));
                range.applyStyle(style, flag);
                style.setNumber(5);
                wsInforme.getCells().get("G" + ((list3 + 17) + list)).setStyle(style);
                wsInforme.getCells().merge(((list3 + 16) + list),6,1,2);
                style = new Style();
                //TOTAL RECAUDADO + EJECUTADO
                wsInforme.getCells().get("C" + ((list3 + 18) + list)).setValue("TOTAL RECAUDADO + EJECUTADO");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getFont().setColor(Color.getWhite());
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("C"+((list3 + 18) + list)+":F" + ((list3 + 18) + list));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge(((list3 + 17) + list),2,1,4);
                style = new Style();
                //CELDA -> RECAUDADO + EJECUTADO
                wsInforme.getCells().get("G" + ((list3 + 18) + list)).setFormula("=G" + ((list3 + 17) + list) + "+G" + ((list3 + 16) + list));
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(146,208,80));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("G"+((list3 + 18) + list)+":H" + ((list3 + 18) + list));
                range.applyStyle(style, flag);
                style.setNumber(5);
                wsInforme.getCells().get("G" + ((list3 + 18) + list)).setStyle(style);
                wsInforme.getCells().merge(((list3 + 17) + list),6,1,2);
                style = new Style();
                //OBSERVACION
                wsInforme.getCells().get("M" + (list3 + 10)).setValue("OBSERVACIÓN: ");
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                range = wsInforme.getCells().createRange("M"+(list3 + 10)+":N" + ((list3 + 12) + dataPorcionXreinstalaciones.get(0).size()));
                range.applyStyle(style, flag);
                wsInforme.getCells().merge((list3 + 9),12,(dataPorcionXreinstalaciones.get(0).size()+3),2);
                style = new Style();

                //PORCIONES
                list = 0;
                value = 0;
                for (c = 'O'; c <= 'P'; c++) {
                    style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                    style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                    style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                    style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                    style.setHorizontalAlignment(TextAlignmentType.CENTER);

                    if (value == 0) {
                        wsInforme.getCells().get("" + c + "" + ((list3 + 12)+list)).setValue(dataPorcionXreinstalaciones.get(value).get(list));
                    } else {
                        wsInforme.getCells().get("" + c + "" + ((list3 + 12)+list)).setValue(Integer.parseInt(dataPorcionXreinstalaciones.get(value).get(list)));
                    }

                    wsInforme.getCells().get("" + c + "" + ((list3 + 12)+list)).setStyle(style);

                    if (value < 1) {
                        value++;
                    }

                    if (list < (dataPorcionXreinstalaciones.get(0).size()-1) && c == 'P') {
                        c = 'N';
                        value = 0;
                        list++;
                    }
                    style = new Style();
                }

                //TOTAL RECAUDADO + EJECUTADO
                wsInforme.getCells().get("O" + ((list3 + 13) + list)).setValue("TOTAL");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.getFont().setColor(Color.getWhite());
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                wsInforme.getCells().get("O" + ((list3 + 13) + list)).setStyle(style);
                style = new Style();
                //CELDA -> RECAUDADO + EJECUTADO
                wsInforme.getCells().get("P" + ((list3 + 13) + list)).setFormula("=SUM(P"+(list3 + 12)+":P"+((list3 + 12)+list)+")");
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(146,208,80));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                wsInforme.getCells().get("P" + ((list3 + 13) + list)).setStyle(style);
                style = new Style();


                //TITULO
                wsInforme.getCells().get("B2").setValue("INFORME DUNNING " + fileMes + " " + anio);
                style.getFont().setSize(28);
                style.getFont().setBold(true);
                style.setForegroundColor(Color.fromArgb(0,176,240));
                style.setPattern(BackgroundType.SOLID);
                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.MEDIUM);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.MEDIUM);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.MEDIUM);
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.MEDIUM);
                range = wsInforme.getCells().createRange("B2:Q2");
                range.applyStyle(style, flag);
                wsInforme.getCells().merge(1,1,1,16);
                style = new Style();

                //BORDE SUPERIOR
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.DOUBLE);
                range = wsInforme.getCells().createRange("C4:P4");
                range.applyStyle(style, flag);
                style = new Style();

                //ESQUINA SUPERIOR IZQUIERDA
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.DOUBLE);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.DOUBLE);
                wsInforme.getCells().get("B4").setStyle(style);
                style = new Style();

                //ESQUINA SUPERIOR DERECHA
                style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.DOUBLE);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.DOUBLE);
                wsInforme.getCells().get("Q4").setStyle(style);
                style = new Style();

                //BORDE IZQUIERDO
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.DOUBLE);
                range = wsInforme.getCells().createRange("B5:B" + ((list3 + 13) + list));
                range.applyStyle(style, flag);
                style = new Style();

                //BORDE DERECHO
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.DOUBLE);
                range = wsInforme.getCells().createRange("Q5:Q" + ((list3 + 13) + list));
                range.applyStyle(style, flag);
                style = new Style();

                //ESQUINA INFERIOR IZQUIERDA
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.DOUBLE);
                style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.DOUBLE);
                wsInforme.getCells().get("B" + ((list3 + 14) + list)).setStyle(style);
                style = new Style();

                //ESQUINA INFERIOR DERECHA
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.DOUBLE);
                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.DOUBLE);
                wsInforme.getCells().get("Q" + ((list3 + 14) + list)).setStyle(style);
                style = new Style();

                //BORDE INFERIOR
                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.DOUBLE);
                range = wsInforme.getCells().createRange("C" + ((list3 + 14) + list) + ":P" + ((list3 + 14) + list));
                range.applyStyle(style, flag);
                style = new Style();

                wb.save("files\\" + mes + ". Informe " + fileMes + "-" + anio + " Dunning.xlsx");
                typeExcel = "INFORME GENERADO CORRECTAMENTE.";
            } catch (Exception e) {
                System.out.println(e);
                codeAlert = 2;
            }

        } catch (Exception e) {
            codeAlert = 1;
        }
    }

    public void excelHistoric() {
        try {
            String[] list = {"IMPRESION","SyT","EXCLUIDAS","REINSTALACIONES"};

            DateFormat date = new SimpleDateFormat("yyyy-MM-dd");

            Calendar calendar = Calendar.getInstance();
            String fecha = date.format(calendar.getTime());

            conexion database = new conexion();
            Connection con = database.conectarSQL();

            int impresiones = 0;
            int syt = 0;
            int excluidas = 0;
            int reinstalaciones = 0;

            for (int i = 0; i < list.length; i++) {
                PreparedStatement ps = con.prepareStatement("SELECT * FROM " + list[i] + " WHERE (f_cierre <= '" + fecha + "') ORDER BY f_cierre");
                ResultSet rs = ps.executeQuery();

                File historico = new File("files\\HISTORICO\\"+list[i]+".csv");
                PrintWriter write = new PrintWriter(historico);

                if (i == 0) {
                    write.print("cuenta_contrato,pagos,porcion,tipo_solicitud,fecha_programado,direccion,f_ejecutado,f_cierre,aviso\n");
                    while (rs.next()) {
                        impresiones++;
                        write.print(rs.getString("cuenta_contrato") + "," + rs.getString("pagos") + "," + rs.getString("porcion") + "," + rs.getString("tipo_solicitud") + "," + rs.getString("fecha") + "," + rs.getString("direccion") + "," + rs.getString("f_ejecutado") + "," + rs.getString("f_cierre") + "," + rs.getString("aviso") + "\n");
                    }
                } else if (i == 1) {
                    write.print("cuenta_contrato,pagos,porcion,tipo_solicitud,fecha_programado,resultado,direccion,f_ejecutado,f_cierre,aviso\n");
                    while (rs.next()) {
                        syt++;
                        write.print(rs.getString("cuenta_contrato") + "," + rs.getString("pagos") + "," + rs.getString("porcion") + "," + rs.getString("tipo_solicitud") + "," + rs.getString("fecha") + "," + rs.getString("resultado") + "," + rs.getString("direccion") + "," + rs.getString("f_ejecutado") + "," + rs.getString("f_cierre") + "," + rs.getString("aviso") + "\n");
                    }
                } else if (i == 2) {
                    write.print("cuenta_contrato,pagos,porcion,tipo_solicitud,fecha_programado,direccion,f_ejecutado,f_cierre,aviso\n");
                    while (rs.next()) {
                        excluidas++;
                        write.print(rs.getString("cuenta_contrato") + "," + rs.getString("pagos") + "," + rs.getString("porcion") + "," + rs.getString("tipo_solicitud") + "," + rs.getString("fecha") + "," + rs.getString("direccion") + "," + rs.getString("f_ejecutado") + "," + rs.getString("f_cierre") + "," + rs.getString("aviso") + "\n");
                    }
                } else if (i == 3) {
                    write.print("cuenta_contrato,porcion,tipo_solicitud,fecha_programado,resultado,direccion,f_cierre,aviso\n");
                    while (rs.next()) {
                        reinstalaciones++;
                        write.print(rs.getString("cuenta_contrato") + "," + rs.getString("porcion") + "," + rs.getString("tipo_solicitud") + "," + rs.getString("fecha") + "," + rs.getString("resultado") + "," + rs.getString("direccion") + "," + rs.getString("f_cierre") + "," + rs.getString("aviso") + "\n");
                    }
                }
                write.close();
            }

            Workbook wbIMPRESION = new Workbook("files\\HISTORICO\\IMPRESION.csv");
            Workbook wbSyT = new Workbook("files\\HISTORICO\\SyT.csv");
            Workbook wbEXCLUIDAS = new Workbook("files\\HISTORICO\\EXCLUIDAS.csv");
            Workbook wbREINSTALACIONES = new Workbook("files\\HISTORICO\\REINSTALACIONES.csv");

            Style style = new Style();
            style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
            style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
            style.setHorizontalAlignment(TextAlignmentType.CENTER);
            wbIMPRESION.getWorksheets().get(0).getCells().createRange("A1:I" + (impresiones+1)).setStyle(style);
            wbSyT.getWorksheets().get(0).getCells().createRange("A1:J" + (syt+1)).setStyle(style);
            wbEXCLUIDAS.getWorksheets().get(0).getCells().createRange("A1:I" + (excluidas+1)).setStyle(style);
            wbREINSTALACIONES.getWorksheets().get(0).getCells().createRange("A1:H" + (reinstalaciones+1)).setStyle(style);
            style.setNumber(5);
            wbIMPRESION.getWorksheets().get(0).getCells().createRange("B2:B" + (impresiones+1)).setStyle(style);
            wbSyT.getWorksheets().get(0).getCells().createRange("B2:B" + (syt+1)).setStyle(style);
            wbEXCLUIDAS.getWorksheets().get(0).getCells().createRange("B2:B" + (excluidas+1)).setStyle(style);

            style.setNumber(14);
            wbIMPRESION.getWorksheets().get(0).getCells().createRange("E2:E" + (impresiones+1)).setStyle(style);
            wbIMPRESION.getWorksheets().get(0).getCells().createRange("G2:G" + (impresiones+1)).setStyle(style);
            wbIMPRESION.getWorksheets().get(0).getCells().createRange("H2:H" + (impresiones+1)).setStyle(style);
            wbSyT.getWorksheets().get(0).getCells().createRange("E2:E" + (syt+1)).setStyle(style);
            wbSyT.getWorksheets().get(0).getCells().createRange("H2:H" + (syt+1)).setStyle(style);
            wbSyT.getWorksheets().get(0).getCells().createRange("I2:I" + (syt+1)).setStyle(style);
            wbEXCLUIDAS.getWorksheets().get(0).getCells().createRange("E2:E" + (excluidas+1)).setStyle(style);
            wbEXCLUIDAS.getWorksheets().get(0).getCells().createRange("G2:G" + (excluidas+1)).setStyle(style);
            wbEXCLUIDAS.getWorksheets().get(0).getCells().createRange("H2:H" + (excluidas+1)).setStyle(style);
            wbREINSTALACIONES.getWorksheets().get(0).getCells().createRange("D2:D" + (reinstalaciones+1)).setStyle(style);
            wbREINSTALACIONES.getWorksheets().get(0).getCells().createRange("G2:G" + (reinstalaciones+1)).setStyle(style);

            style.setForegroundColor(Color.fromArgb(255, 255, 102));
            style.setPattern(BackgroundType.SOLID);

            wbIMPRESION.getWorksheets().get(0).getCells().createRange("A1:I1").setStyle(style);
            Cells cells = wbIMPRESION.getWorksheets().get(0).getCells();
            cells.setColumnWidth(0, 13.43); //A
            cells.setColumnWidth(1, 14); //B
            cells.setColumnWidth(2, 6.29); //C
            cells.setColumnWidth(3, 14.86); //D
            cells.setColumnWidth(4, 15.43); //E
            cells.setColumnWidth(5, 40); //F
            cells.setColumnWidth(6, 10); //G
            cells.setColumnWidth(7, 10); //H
            cells.setColumnWidth(8, 10); //I

            wbSyT.getWorksheets().get(0).getCells().createRange("A1:J1").setStyle(style);
            cells = wbSyT.getWorksheets().get(0).getCells();
            cells.setColumnWidth(0, 13.43); //A
            cells.setColumnWidth(1, 14); //B
            cells.setColumnWidth(2, 6.29); //C
            cells.setColumnWidth(3, 14.86); //D
            cells.setColumnWidth(4, 15.43); //E
            cells.setColumnWidth(5, 8); //F
            cells.setColumnWidth(6, 40); //G
            cells.setColumnWidth(7, 10); //H
            cells.setColumnWidth(8, 10); //I
            cells.setColumnWidth(9, 10); //J

            wbEXCLUIDAS.getWorksheets().get(0).getCells().createRange("A1:I1").setStyle(style);
            cells = wbEXCLUIDAS.getWorksheets().get(0).getCells();
            cells.setColumnWidth(0, 13.43); //A
            cells.setColumnWidth(1, 14); //B
            cells.setColumnWidth(2, 6.29); //C
            cells.setColumnWidth(3, 14.86); //D
            cells.setColumnWidth(4, 15.43); //E
            cells.setColumnWidth(5, 40); //F
            cells.setColumnWidth(6, 10); //G
            cells.setColumnWidth(7, 10); //H
            cells.setColumnWidth(8, 10); //I

            wbREINSTALACIONES.getWorksheets().get(0).getCells().createRange("A1:H1").setStyle(style);
            cells = wbREINSTALACIONES.getWorksheets().get(0).getCells();
            cells.setColumnWidth(0, 13.43); //A
            cells.setColumnWidth(1, 6); //B
            cells.setColumnWidth(2, 15); //C
            cells.setColumnWidth(3, 14.86); //D
            cells.setColumnWidth(4, 8); //E
            cells.setColumnWidth(5, 40); //F
            cells.setColumnWidth(6, 10); //G
            cells.setColumnWidth(7, 10); //H

            wbIMPRESION.save("files\\HISTORICO\\IMPRESION.xlsx");
            wbSyT.save("files\\HISTORICO\\SyT.xlsx");
            wbEXCLUIDAS.save("files\\HISTORICO\\EXCLUIDAS.xlsx");
            wbREINSTALACIONES.save("files\\HISTORICO\\REINSTALACIONES.xlsx");

            Workbook wb = new Workbook();
            wb.combine(wbIMPRESION);
            wb.combine(wbSyT);
            wb.combine(wbEXCLUIDAS);
            wb.combine(wbREINSTALACIONES);
            wb.getWorksheets().removeAt(0);

            DateFormat dateFile = new SimpleDateFormat("dd-MM-yyyy");
            Date fechaDate = date.parse(fecha);
            String fechaFile = dateFile.format(fechaDate);

            try {
                wb.save("files\\HISTORICO\\Historico STR " + fechaFile + ".xlsx");
                typeExcel = "HISTORICO GENERADO CORRECTAMENTE.";
            } catch (Exception e) {
                codeAlert = 2;
                System.out.println(e);
            }

            new File("files\\HISTORICO\\IMPRESION.csv").delete();
            new File("files\\HISTORICO\\SyT.csv").delete();
            new File("files\\HISTORICO\\EXCLUIDAS.csv").delete();
            new File("files\\HISTORICO\\REINSTALACIONES.csv").delete();
            new File("files\\HISTORICO\\IMPRESION.xlsx").delete();
            new File("files\\HISTORICO\\SyT.xlsx").delete();
            new File("files\\HISTORICO\\EXCLUIDAS.xlsx").delete();
            new File("files\\HISTORICO\\REINSTALACIONES.xlsx").delete();
        } catch(Exception e) {
            System.out.println(e);
        }
    }

    String fileError = "";
    public void excelDeudores(Workbook wbHistoric, Workbook wbCC) {
        Worksheet ws1 = wbHistoric.getWorksheets().get(1);
        List<Integer> cuenta_contrato = new ArrayList<>();
        List<Integer> pagos = new ArrayList<>();
        List<String> porcion = new ArrayList<>();
        List<String> tipo_solicitud = new ArrayList<>();
        List<String> fecha = new ArrayList<>();
        List<Integer> resultado = new ArrayList<>();
        List<String> direccion = new ArrayList<>();
        List<String> f_ejecutado = new ArrayList<>();
        List<String> f_cierre = new ArrayList<>();
        List<String> aviso = new ArrayList<>();

        if (ws1.getName().equals("SyT") && ws1.getCells().getMaxDataColumn() == 9) {
            for (int i = 1; i < (ws1.getCells().getLastDataRow(0)+1); i++) {
                if (ws1.getCells().get(i, 5).getIntValue() == 1) {
                    try {
                        cuenta_contrato.add(ws1.getCells().get(i, 0).getIntValue());
                        pagos.add(ws1.getCells().get(i, 1).getIntValue());
                    } catch (Exception e) {
                        codeAlert = 4;
                    }

                    if (ws1.getCells().get(i, 2).getValue() != null && ws1.getCells().get(i, 2).getValue() != "") {
                        porcion.add(ws1.getCells().get(i, 2).getStringValue());
                    } else {
                        codeAlert = 4;
                    }

                    if (ws1.getCells().get(i, 3).getValue() != null && ws1.getCells().get(i, 3).getValue() != "") {
                        tipo_solicitud.add(ws1.getCells().get(i, 3).getStringValue());
                    } else {
                        codeAlert = 4;
                    }

                    DateFormat dateFormat = new SimpleDateFormat("d/MM/yyyy");
                    SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");

                    String validFecha = ws1.getCells().get(i, 4).getStringValue();
                    try {
                        Date date = dateFormat.parse(validFecha);
                        validFecha = simpleDateFormat.format(date);
                        if (validFecha.length() == 10) {
                            fecha.add(validFecha);
                        } else {
                            codeAlert = 4;
                        }
                    } catch (Exception e) {
                        codeAlert = 4;
                    }

                    try {
                        resultado.add(ws1.getCells().get(i, 5).getIntValue());
                    } catch (Exception e) {
                        codeAlert = 4;
                    }

                    if (ws1.getCells().get(i, 6).getValue() != null && ws1.getCells().get(i, 6).getValue() != "") {
                        String dir = ws1.getCells().get(i, 6).getStringValue();
                        dir.replaceAll(",", "");
                        dir.replaceAll(";", "");
                        dir.replaceAll("\"", "");
                        direccion.add(dir);
                    } else {
                        codeAlert = 4;
                    }

                    validFecha = ws1.getCells().get(i, 7).getStringValue();
                    try {
                        Date date = dateFormat.parse(validFecha);
                        validFecha = simpleDateFormat.format(date);
                        if (validFecha.length() == 10) {
                            f_ejecutado.add(validFecha);
                        } else {
                            codeAlert = 4;
                        }
                    } catch (Exception e) {
                        codeAlert = 4;
                    }

                    validFecha = ws1.getCells().get(i, 8).getStringValue();
                    try {
                        Date date = dateFormat.parse(validFecha);
                        validFecha = simpleDateFormat.format(date);
                        if (validFecha.length() == 10) {
                            f_cierre.add(validFecha);
                        } else {
                            codeAlert = 4;
                        }
                    } catch (Exception e) {
                        codeAlert = 4;
                    }

                    if (ws1.getCells().get(i, 9).getValue() != null && ws1.getCells().get(i, 9).getValue() != "") {
                        aviso.add(ws1.getCells().get(i, 9).getStringValue());
                    } else {
                        codeAlert = 4;
                    }
                }
            }

            if (codeAlert != 4) {
                File tempHistorico = new File("files\\DEUDORES\\tempHIS.csv");

                try {
                    PrintWriter write = new PrintWriter(tempHistorico);
                    for (int i = 0; i < cuenta_contrato.size(); i++) {
                        String lineData = cuenta_contrato.get(i) + " , " + pagos.get(i) + " , " + porcion.get(i) + " , " + tipo_solicitud.get(i) + " , " + fecha.get(i) + " , " + resultado.get(i) + " , " + direccion.get(i) + " , " + f_ejecutado.get(i) + " , " + f_cierre.get(i) + " , " + aviso.get(i);
                        write.println(lineData);
                    }
                    write.close();

                    File uploadFiles = new File("tools\\shell\\files.txt");
                    write = new PrintWriter(uploadFiles);
                    write.println(".mode csv");
                    write.println(".open '" + new File("tools\\db\\database.db").getAbsolutePath() + "'");
                    write.println(".import '" + tempHistorico.getAbsolutePath() + "' tempHISTORICO");
                    write.println(".shell del '" + tempHistorico.getAbsolutePath() + "'");
                    write.println(".exit");
                    write.close();

                    Process p = Runtime.getRuntime().exec("cmd /c cd " + new File("tools\\shell").getPath() + " && upload.cmd");
                    p.getErrorStream().close();
                    p.waitFor();

                } catch (Exception e) {
                    System.out.println(e);
                }
            }
        } else {
            codeAlert = 3;
        }

        if (codeAlert == 4) {
            fileError += "HISTORICO";
        }

        Worksheet ws2 = wbCC.getWorksheets().get(0);

        List<Integer> cuenta_contrato_CC = new ArrayList<>();
        List<Integer> interloc = new ArrayList<>();

        if (ws2.getCells().getMaxDataColumn() == 10) {
            for (int i = 1; i < (ws2.getCells().getLastDataRow(0)+1); i++) {
                try {
                    cuenta_contrato_CC.add(ws2.getCells().get(i,0).getIntValue());
                    interloc.add(ws2.getCells().get(i,1).getIntValue());
                } catch (Exception e) {
                    codeAlert = 4;
                }
            }

            if (codeAlert != 4) {
                File tempUSUARIOS = new File("files\\DEUDORES\\tempUSUARIOS.csv");
                try {
                    PrintWriter write = new PrintWriter(tempUSUARIOS);
                    for (int i = 0; i < cuenta_contrato_CC.size(); i++) {
                        write.print(cuenta_contrato_CC.get(i) + "," + interloc.get(i) + "\n");
                    }
                    write.close();

                    File uploadFiles = new File("tools\\shell\\files.txt");
                    write = new PrintWriter(uploadFiles);
                    write.println(".mode csv");
                    write.println(".open '" + new File("tools\\db\\database.db").getAbsolutePath() + "'");
                    write.println(".import '" + tempUSUARIOS.getAbsolutePath() + "' tempUSUARIOS");
                    write.println(".shell del '" + tempUSUARIOS.getAbsolutePath() + "'");
                    write.println(".exit");
                    write.close();

                    Process p = Runtime.getRuntime().exec("cmd /c cd " + new File("tools\\shell").getPath() + " && upload.cmd");
                    p.getErrorStream().close();
                    p.waitFor();

                } catch (Exception e) {
                    System.out.println(e);
                }
            }

            if (codeAlert == 4 && fileError != "") {
                fileError += " Y ";
            }

            if (codeAlert == 4) {
                fileError += "TOTAL DE USUARIOS (CUENTAS CONTRATOS).";
            }

            if (codeAlert != 4) {
                conexion database = new conexion();
                Connection con = database.conectarSQL();
                try {
                    PreparedStatement ps = con.prepareStatement("SELECT tempHISTORICO.cuenta_contrato, tempHISTORICO.pagos, tempHISTORICO.porcion, tempHISTORICO.tipo_solicitud, tempHISTORICO.fecha_programado, tempHISTORICO.resultado, tempHISTORICO.direccion, tempHISTORICO.f_ejecutado, tempHISTORICO.f_cierre, tempHISTORICO.aviso, tempUSUARIOS.interloc FROM tempHISTORICO LEFT JOIN tempUSUARIOS ON tempHistorico.cuenta_contrato = tempUSUARIOS.cuenta_contrato");
                    ResultSet rs = ps.executeQuery();

                    Workbook wbDeudores = new Workbook();

                    String[] estructura = {"cuenta_contrato","pagos","porcion","tipo de solicitud", "fecha de programacion", "resultado", "direccion", "f_ejecutado", "f_cierre", "aviso", "interlocutor", "descripcion"};

                    int j = 1;
                    while (rs.next()) {
                        for (int i = 0; i < estructura.length; i++) {
                            Style style = new Style();
                            style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                            style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                            style.setHorizontalAlignment(TextAlignmentType.CENTER);

                            if (j == 1) {
                                wbDeudores.getWorksheets().get(0).getCells().get(0, i).setValue(estructura[i]);
                                style.setForegroundColor(Color.fromArgb(255,255,102));
                                style.setPattern(BackgroundType.SOLID);
                                wbDeudores.getWorksheets().get(0).getCells().get(0, i).setStyle(style);
                                style = new Style();
                                style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
                                style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
                                style.setHorizontalAlignment(TextAlignmentType.CENTER);
                            }

                            if (i == 0) {
                                wbDeudores.getWorksheets().get(0).getCells().get(j, i).setValue(Integer.parseInt(rs.getString("cuenta_contrato")));
                            } else if (i == 1) {
                                wbDeudores.getWorksheets().get(0).getCells().get(j, i).setValue(Integer.parseInt(rs.getString("pagos")));
                            } else if (i == 2) {
                                wbDeudores.getWorksheets().get(0).getCells().get(j, i).setValue(rs.getString("porcion"));
                            } else if (i == 3) {
                                wbDeudores.getWorksheets().get(0).getCells().get(j, i).setValue(rs.getString("tipo_solicitud"));
                            } else if (i == 4) {
                                wbDeudores.getWorksheets().get(0).getCells().get(j, i).setValue(rs.getString("fecha_programado"));
                            } else if (i == 5) {
                                wbDeudores.getWorksheets().get(0).getCells().get(j, i).setValue(Integer.parseInt(rs.getString("resultado")));
                            } else if (i == 6) {
                                wbDeudores.getWorksheets().get(0).getCells().get(j, i).setValue(rs.getString("direccion"));
                            } else if (i == 7) {
                                wbDeudores.getWorksheets().get(0).getCells().get(j, i).setValue(rs.getString("f_ejecutado"));
                            } else if (i == 8) {
                                wbDeudores.getWorksheets().get(0).getCells().get(j, i).setValue(rs.getString("f_cierre"));
                            } else if (i == 9) {
                                wbDeudores.getWorksheets().get(0).getCells().get(j, i).setValue(rs.getString("aviso"));
                            } else if (i == 10) {
                                if (rs.getString("interloc") == null) {
                                    wbDeudores.getWorksheets().get(0).getCells().get(j, i).setValue("#N/A");
                                } else {
                                    wbDeudores.getWorksheets().get(0).getCells().get(j, i).setValue(Integer.parseInt(rs.getString("interloc")));
                                }

                            } else if (i == 11) {
                                if (Integer.parseInt(rs.getString("pagos")) < 200000) {
                                    wbDeudores.getWorksheets().get(0).getCells().get(j, i).setValue("MENOR DE 200 MIL");
                                    style.setForegroundColor(Color.fromArgb(255,255,0));
                                    style.setPattern(BackgroundType.SOLID);
                                    wbDeudores.getWorksheets().get(0).getCells().get(j, i).setStyle(style);
                                } else if (Integer.parseInt(rs.getString("pagos")) >= 200000 && Integer.parseInt(rs.getString("pagos")) < 600000) {
                                    wbDeudores.getWorksheets().get(0).getCells().get(j, i).setValue("ENTRE 200 Y 600 MIL");
                                    style.setForegroundColor(Color.fromArgb(255,192,0));
                                    style.setPattern(BackgroundType.SOLID);
                                    wbDeudores.getWorksheets().get(0).getCells().get(j, i).setStyle(style);
                                } else if (Integer.parseInt(rs.getString("pagos")) >= 600000 && Integer.parseInt(rs.getString("pagos")) < 1000000) {
                                    wbDeudores.getWorksheets().get(0).getCells().get(j, i).setValue("ENTRE 600 MIL Y 1 MILLON");
                                    style.setForegroundColor(Color.fromArgb(197,83,13));
                                    style.setPattern(BackgroundType.SOLID);
                                    wbDeudores.getWorksheets().get(0).getCells().get(j, i).setStyle(style);
                                } else if (Integer.parseInt(rs.getString("pagos")) >= 1000000) {
                                    wbDeudores.getWorksheets().get(0).getCells().get(j, i).setValue("MAYOR A 1 MILLON");
                                    style.setForegroundColor(Color.fromArgb(255,0,0));
                                    style.setPattern(BackgroundType.SOLID);
                                    wbDeudores.getWorksheets().get(0).getCells().get(j, i).setStyle(style);
                                }
                            }

                            wbDeudores.getWorksheets().get(0).getCells().get(j, i).setStyle(style);
                        }
                        j++;
                    }

                    Cells cells = wbDeudores.getWorksheets().get(0).getCells();
                    cells.setColumnWidth(0, 13.43); //A
                    cells.setColumnWidth(1, 10); //B
                    cells.setColumnWidth(2, 6.29); //C
                    cells.setColumnWidth(3, 15); //D
                    cells.setColumnWidth(4, 15.43); //E
                    cells.setColumnWidth(5, 8); //F
                    cells.setColumnWidth(6, 40); //G
                    cells.setColumnWidth(7, 10); //H
                    cells.setColumnWidth(8, 10); //I
                    cells.setColumnWidth(9, 10); //J
                    cells.setColumnWidth(10, 13.57); //K
                    cells.setColumnWidth(11, 24.14); //L

                    DateFormat date = new SimpleDateFormat("dd-MM-yyyy");
                    Calendar calendar = Calendar.getInstance();
                    String fechaActual = date.format(calendar.getTime());
                    try {
                        wbDeudores.getWorksheets().get(0).setName("DEUDORES " + fechaActual);
                        wbDeudores.save("files\\DEUDORES\\Deudores " + fechaActual + ".xlsx");

                        ps = con.prepareStatement("DELETE FROM tempHISTORICO");
                        ps.executeUpdate();
                        ps = con.prepareStatement("DELETE FROM tempUSUARIOS");
                        ps.executeUpdate();

                        typeExcel = "LISTA DE DEUDORES GENERADO CORRECTAMENTE.";
                    } catch (Exception e) {
                        codeAlert = 2;
                    }
                } catch (Exception e) {
                    System.out.println(e);
                }
            }
        } else {
            codeAlert = 3;
        }
    }

    public void generate(Stage initStage, String month, String year, int typeExcelButton, File historicFile, File ccFile) {
        if (typeExcelButton == 1) {
            new Thread(() -> {excelInformes(month, year);}).run();
        } else if (typeExcelButton == 2) {
            new Thread (() -> {excelHistoric();}).run();
        } else if (typeExcelButton == 3) {
            try {
                Workbook wbHistoric = new Workbook(historicFile.getAbsolutePath());
                Workbook wbCC = new Workbook(ccFile.getAbsolutePath());
                new Thread (() -> {excelDeudores(wbHistoric, wbCC);}).run();
            } catch (Exception e) {
                codeAlert = 4;
                System.out.println(e);
            }
        }
        Platform.runLater(new Runnable() {
            @Override
            public void run() {
                new loading(initStage);
                Stage primaryStage = new Stage();
                new app().start(primaryStage);

                if (codeAlert == 0) {
                    Alert alert = new Alert(Alert.AlertType.INFORMATION);
                    alert.setHeaderText(null);
                    alert.setTitle("Success");
                    alert.setContentText(typeExcel);
                    alert.showAndWait();

                    if (typeExcelButton == 1) {
                        try {
                            File openInforme = new File("files");
                            Runtime.getRuntime().exec("cmd /c start " + openInforme.getAbsolutePath() + " && exit");
                        } catch (Exception e) {
                            System.out.println(e);
                        }
                    } else if (typeExcelButton == 2) {
                        try {
                            File openInforme = new File("files\\HISTORICO");
                            Runtime.getRuntime().exec("cmd /c start " + openInforme.getAbsolutePath() + " && exit");
                        } catch (Exception e) {
                            System.out.println(e);
                        }
                    } else if (typeExcelButton == 3) {
                        try {
                            File openInforme = new File("files\\DEUDORES");
                            Runtime.getRuntime().exec("cmd /c start " + openInforme.getAbsolutePath() + " && exit");
                        } catch (Exception e) {
                            System.out.println(e);
                        }
                    }
                } else if (codeAlert == 1) {
                    Alert alert = new Alert(Alert.AlertType.INFORMATION);
                    alert.setHeaderText(null);
                    alert.setTitle("Error");
                    alert.setContentText("NO HAY DATOS DE LA RESPECTIVA FECHA.");
                    alert.showAndWait();
                } else if (codeAlert == 2) {
                    Alert alert = new Alert(Alert.AlertType.INFORMATION);
                    alert.setHeaderText(null);
                    alert.setTitle("Error");
                    alert.setContentText("EL ARCHIVO A GENERAR SE ENCUENTRA ABIERTO, CIERRO PARA CONTINUAR.");
                    alert.showAndWait();
                } else if (codeAlert == 3) {
                    Alert alert = new Alert(Alert.AlertType.INFORMATION);
                    alert.setHeaderText(null);
                    alert.setTitle("Error");
                    alert.setContentText("VERIFIQUE QUE SEAN LOS ARCHIVOS CORRESPONDIENTES.");
                    alert.showAndWait();
                } else if (codeAlert == 4) {
                    Alert alert = new Alert(Alert.AlertType.INFORMATION);
                    alert.setHeaderText(null);
                    alert.setTitle("Error");
                    alert.setContentText("VERIFIQUE LOS DATOS DEL ARCHIVO "+fileError+".");
                    alert.showAndWait();
                }
            }
        });
    }
}
