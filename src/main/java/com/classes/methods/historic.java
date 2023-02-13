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
import java.util.Calendar;
import java.util.Date;

public class historic {
    int codeAlert = 0;

    public void excel() {
        try {
            String[] list = {"IMPRESION","SyT","EXCLUIDAS","REINSTALACIONES"};

            DateFormat date = new SimpleDateFormat("yyyy-MM-dd");

            Calendar calendar = Calendar.getInstance();
            String fecha = date.format(calendar.getTime());

            conexion database = new conexion();
            Connection con = database.conectarSQL();

            for (int i = 0; i < list.length; i++) {
                PreparedStatement ps = con.prepareStatement("SELECT * FROM " + list[i] + " WHERE (f_cierre <= '" + fecha + "') ORDER BY f_cierre");
                ResultSet rs = ps.executeQuery();

                File historico = new File("files\\HISTORICO\\"+list[i]+".csv");
                PrintWriter write = new PrintWriter(historico);

                if (i == 0) {
                    write.print("cuenta_contrato,pagos,porcion,tipo_solicitud,fecha_programado,direccion,f_ejecutado,f_cierre,aviso\n");
                    while (rs.next()) {
                        write.print(rs.getString("cuenta_contrato") + "," + rs.getString("pagos") + "," + rs.getString("porcion") + "," + rs.getString("tipo_solicitud") + "," + rs.getString("fecha") + "," + rs.getString("direccion") + "," + rs.getString("f_ejecutado") + "," + rs.getString("f_cierre") + "," + rs.getString("aviso") + "\n");
                    }
                } else if (i == 1) {
                    write.print("cuenta_contrato,pagos,porcion,tipo_solicitud,fecha_programado,resultado,direccion,f_ejecutado,f_cierre,aviso\n");
                    while (rs.next()) {
                        write.print(rs.getString("cuenta_contrato") + "," + rs.getString("pagos") + "," + rs.getString("porcion") + "," + rs.getString("tipo_solicitud") + "," + rs.getString("fecha") + "," + rs.getString("resultado") + "," + rs.getString("direccion") + "," + rs.getString("f_ejecutado") + "," + rs.getString("f_cierre") + "," + rs.getString("aviso") + "\n");
                    }
                } else if (i == 2) {
                    write.print("cuenta_contrato,pagos,porcion,tipo_solicitud,fecha_programado,direccion,f_ejecutado,f_cierre,aviso\n");
                    while (rs.next()) {
                        write.print(rs.getString("cuenta_contrato") + "," + rs.getString("pagos") + "," + rs.getString("porcion") + "," + rs.getString("tipo_solicitud") + "," + rs.getString("fecha") + "," + rs.getString("direccion") + "," + rs.getString("f_ejecutado") + "," + rs.getString("f_cierre") + "," + rs.getString("aviso") + "\n");
                    }
                } else if (i == 3) {
                    write.print("cuenta_contrato,porcion,tipo_solicitud,fecha_programado,resultado,direccion,f_cierre,aviso\n");
                    while (rs.next()) {
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
            style.setForegroundColor(Color.fromArgb(255, 255, 102));
            style.setPattern(BackgroundType.SOLID);

            wbIMPRESION.getWorksheets().get(0).getCells().createRange("A1:I1").setStyle(style);
            Cells cells = wbIMPRESION.getWorksheets().get(0).getCells();
            cells.setColumnWidth(0, 13.43); //A
            cells.setColumnWidth(1, 9.43); //B
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
            cells.setColumnWidth(1, 9.43); //B
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
            cells.setColumnWidth(1, 9.43); //B
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
            } catch (Exception e) {
                codeAlert = 1;
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

            if (codeAlert == 0) {
                File openHistorico = new File("files\\HISTORICO");
                Runtime.getRuntime().exec("cmd /c start " + openHistorico.getAbsolutePath() + " && exit");
            }
        } catch(Exception e) {
            System.out.println(e);
        }
    }

    public void historic (Stage initStage) {
        new Thread (() -> {excel();}).run();
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
                    alert.setContentText("HISTORICO GENERADO CORRECTAMENTE.");
                    alert.showAndWait();
                } else if (codeAlert == 1) {
                    Alert alert = new Alert(Alert.AlertType.INFORMATION);
                    alert.setHeaderText(null);
                    alert.setTitle("Error");
                    alert.setContentText("EL ARCHIVO A GENERAR SE ENCUENTRA ABIERTO, CIERRO PARA CONTINUAR.");
                    alert.showAndWait();
                }

            }
        });
    }
}

