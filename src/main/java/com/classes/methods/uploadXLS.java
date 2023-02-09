package com.classes.methods;

import com.app.app;
import com.aspose.cells.*;
import com.classes.connection.conexion;
import javafx.application.Platform;
import javafx.scene.control.Alert;
import javafx.scene.control.TextField;
import javafx.stage.Stage;
import java.io.File;
import java.io.PrintWriter;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.List;
import java.util.concurrent.atomic.AtomicBoolean;

public class uploadXLS {

    boolean isError = false;
    int codeError = 0;
    String typeError = "";

    public void validXLSX(Workbook wbXLSX) {
        if (wbXLSX.getWorksheets().getCount() == 1) {
            List<List> dataWS1 = new ArrayList<>();

            Worksheet ws1 = wbXLSX.getWorksheets().get(0);
            if ((ws1.getName().equals("Reinstalaciones") || ws1.getName().equals("reinstalaciones") || ws1.getName().equals("REINSTALACIONES") || ws1.getName().equals("REINSTALACIÓN") || ws1.getName().equals("REINSTALACION") || ws1.getName().equals("reinstalacion") || ws1.getName().equals("Reinstalación")))
            {
                if (ws1.getCells().getMaxDataColumn()+1 == 18) {
                    if (ws1.getCells().get(0,2).getType() == 5 && ws1.getCells().get(0,4).getType() == 5 && ws1.getCells().get(0,10).getType() == 5 && ws1.getCells().get(0,16).getType() == 5 && ws1.getCells().get(0,17).getType() == 5) {
                        if (
                        (ws1.getCells().getMaxDataRow()+1 == ws1.getCells().getLastDataRow(2)+1) &&
                        (ws1.getCells().getMaxDataRow()+1 == ws1.getCells().getLastDataRow(4)+1) &&
                        (ws1.getCells().getMaxDataRow()+1 == ws1.getCells().getLastDataRow(10)+1) &&
                        (ws1.getCells().getMaxDataRow()+1 == ws1.getCells().getLastDataRow(16)+1) &&
                        (ws1.getCells().getMaxDataRow()+1 == ws1.getCells().getLastDataRow(17)+1)
                        ) {
                            boolean error = false;
                            String[] errorType = {"","","","","", ""};

                            List<String> avisos = new ArrayList<>();
                            List<String> porcion = new ArrayList<>();
                            List<String> tipoSolicitud = new ArrayList<>();
                            List<String> fecha = new ArrayList<>();
                            List<Integer> resultado = new ArrayList<>();
                            List<String> f_cierre = new ArrayList<>();

                            DateFormat dateFormat = new SimpleDateFormat("d/MM/yyyy");
                            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");

                            for (int i = 1; i < (ws1.getCells().getLastDataRow(0)+1); i++) {
                                if (ws1.getCells().get(i,17).getValue() != null && ws1.getCells().get(i,17).getValue() != "" && ws1.getCells().get(i,17).getStringValue().contains("CX")) {
                                    avisos.add(ws1.getCells().get(i,17).getStringValue());
                                } else {
                                    error = true;
                                    errorType[0] = "\n• numero de aviso";
                                }

                                if (ws1.getCells().get(i,2).getValue() != null && ws1.getCells().get(i,2).getValue() != "" && (ws1.getCells().get(i,2).getValue().equals("REINSTALACION"))) {
                                    tipoSolicitud.add(ws1.getCells().get(i,2).getStringValue());
                                } else {
                                    error = true;
                                    errorType[1] = "\n• tipo de solicitud";
                                }

                                if (ws1.getCells().get(i,6).getValue() != null && ws1.getCells().get(i,6).getValue() != "") {
                                    porcion.add(ws1.getCells().get(i,6).getStringValue());
                                } else {
                                    error = true;
                                    errorType[2] = "\n• porcion";
                                }

                                try {
                                    String validFecha = ws1.getCells().get(i,4).getStringValue();
                                    Date date = dateFormat.parse(validFecha);
                                    validFecha = simpleDateFormat.format(date);
                                    if (validFecha.length() == 10) {
                                        fecha.add(validFecha);
                                    } else {
                                        error = true;
                                        errorType[3] = "\n• fecha de programación";
                                    }
                                } catch (Exception e) {
                                    error = true;
                                    errorType[3] = "\n• fecha de programación";
                                }

                                try {
                                    resultado.add(ws1.getCells().get(i,10).getIntValue());
                                } catch (Exception e) {
                                    error = true;
                                    errorType[4] = "\n• resultado";
                                }

                                try {
                                    String validFecha = ws1.getCells().get(i,16).getStringValue();
                                    Date date = dateFormat.parse(validFecha);
                                    validFecha = simpleDateFormat.format(date);
                                    if (validFecha.length() == 10) {
                                        f_cierre.add(validFecha);
                                    } else {
                                        error = true;
                                        errorType[5] = "\n• fecha de cierre";
                                    }
                                } catch (Exception e) {
                                    error = true;
                                    errorType[5] = "\n• fecha de cierre";
                                }
                            }

                            if (error != true) {
                                dataWS1.add(avisos);
                                dataWS1.add(porcion);
                                dataWS1.add(tipoSolicitud);
                                dataWS1.add(resultado);
                                dataWS1.add(fecha);
                                dataWS1.add(f_cierre);

                                File reinstalacionesCSV = new File("files\\data_Reinstalaciones.csv");

                                try {
                                    PrintWriter write = new PrintWriter(reinstalacionesCSV);
                                    for (int j = 0; j < dataWS1.get(0).size(); j++) {
                                        for (int i = 0; i < dataWS1.size(); i++) {
                                            write.print(dataWS1.get(i).get(j));
                                            if (i < (dataWS1.size()-1)) {
                                                write.print(",");
                                            } else if (i == (dataWS1.size()-1)) {
                                                write.print("\n");
                                            }
                                        }
                                    }
                                    write.close();

                                    File uploadFiles = new File("tools\\shell\\files.txt");
                                    write = new PrintWriter(uploadFiles);
                                    write.println(".mode csv");
                                    write.println(".open '" + new File("tools\\db\\database.db").getAbsolutePath() + "'");
                                    write.println(".import '" + reinstalacionesCSV.getAbsolutePath() + "' REINSTALACIONES");
                                    write.println(".shell del '" + reinstalacionesCSV.getAbsolutePath() + "'");
                                    write.println(".exit");
                                    write.close();

                                    Process p = Runtime.getRuntime().exec("cmd /c cd " + new File("tools\\shell").getPath() + " && upload.cmd");
                                    p.getErrorStream().close();
                                    p.waitFor();

                                    String deleteDate = dataWS1.get(4).get(0).toString();
                                    try {
                                        dateFormat = new SimpleDateFormat("yyyy-MM-dd");
                                        Calendar c = Calendar.getInstance();
                                        c.setTime(dateFormat.parse(deleteDate));
                                        c.add(Calendar.YEAR, -1);
                                        deleteDate = dateFormat.format(c.getTime());
                                    } catch (Exception e) {
                                        System.out.println(e);
                                    }

                                    conexion database = new conexion();
                                    Connection con = database.conectarSQL();

                                    PreparedStatement ps = con.prepareStatement("DELETE FROM REINSTALACIONES WHERE (fecha < '"+deleteDate+"');");
                                    ps.executeUpdate();

                                } catch (Exception e) {
                                    System.out.println(e);
                                }


                            } else {
                                typeError += "\n➤ IMPRESIÓN" + errorType[0] + errorType[1] + errorType[2] + errorType[3] + errorType[4] + errorType[5] + "\n";
                                isError = true;
                                codeError = 5;
                            }

                        } else {
                            isError = true;
                            codeError = 4;
                        }
                    } else {
                        codeError = 3;
                        isError = true;
                        typeError += "\n➤ REINSTALACIONES";
                    }
                } else {
                    codeError = 3;
                    isError = true;
                    typeError += "\n➤ REINSTALACIONES";
                }
            } else {
                codeError = 2;
                isError = true;
                typeError = "\n\n1) REINSTALACIONES";
            }

        } else if (wbXLSX.getWorksheets().getCount() == 3) {
            List<List> dataWS1 = new ArrayList<>();
            List<List> dataWS2 = new ArrayList<>();
            List<List> dataWS3 = new ArrayList<>();

            Worksheet ws1 = wbXLSX.getWorksheets().get(0);
            Worksheet ws2 = wbXLSX.getWorksheets().get(1);
            Worksheet ws3 = wbXLSX.getWorksheets().get(2);
            if (
            (ws1.getName().equals("IMPRESIÓN") || ws1.getName().equals("IMPRESION") || ws1.getName().equals("IMPRESIONES") || ws1.getName().equals("impresión") || ws1.getName().equals("impresion") || ws1.getName().equals("impresiones") || ws1.getName().equals("IMPRIMIR") || ws1.getName().equals("imprimir")) &&
            (ws2.getName().equals("SyT") || ws2.getName().equals("syt") || ws2.getName().equals("SYT") || ws2.getName().equals("SYt") || ws2.getName().equals("Syt") || ws2.getName().equals("sYT") || ws2.getName().equals("syT")) &&
            (ws3.getName().equals("EXCLUIDAS") || ws3.getName().equals("EXCLUIDOS") || ws3.getName().equals("excluidas") || ws3.getName().equals("excluidos") || ws3.getName().equals("EXCLUSIONES") || ws3.getName().equals("exclusiones") || ws3.getName().equals("EXCLUSION") || ws3.getName().equals("EXCLUSIÓN") || ws3.getName().equals("exclusion") || ws3.getName().equals("exclusión"))
            ) {
                if (ws1.getCells().getMaxDataColumn()+1 == 17 && ws2.getCells().getMaxDataColumn()+1 == 21 && ws3.getCells().getMaxDataColumn()+1 == 13) {

                    if (
                    (ws1.getCells().get(0,1).getType() == 5 && ws1.getCells().get(0,3).getType() == 5 && ws1.getCells().get(0,5).getType() == 5 && ws1.getCells().get(0,7).getType() == 5 && ws1.getCells().get(0,12).getType() == 5 && ws1.getCells().get(0,13).getType() == 5) &&
                    (ws2.getCells().get(0,1).getType() == 5 && ws2.getCells().get(0,3).getType() == 5 && ws2.getCells().get(0,5).getType() == 5 && ws2.getCells().get(0,7).getType() == 5 && ws2.getCells().get(0,11).getType() == 5 && ws2.getCells().get(0,17).getType() == 5) &&
                    (ws3.getCells().get(0,1).getType() == 5 && ws3.getCells().get(0,3).getType() == 5 && ws3.getCells().get(0,5).getType() == 5 && ws3.getCells().get(0,7).getType() == 5 && ws3.getCells().get(0,11).getType() == 5)
                    ) {
                        if (
                        (ws1.getCells().getMaxDataRow()+1 == ws1.getCells().getLastDataRow(1)+1) &&
                        (ws1.getCells().getMaxDataRow()+1 == ws1.getCells().getLastDataRow(3)+1) &&
                        (ws1.getCells().getMaxDataRow()+1 == ws1.getCells().getLastDataRow(5)+1) &&
                        (ws1.getCells().getMaxDataRow()+1 == ws1.getCells().getLastDataRow(7)+1) &&
                        (ws1.getCells().getMaxDataRow()+1 == ws1.getCells().getLastDataRow(12)+1) &&
                        (ws1.getCells().getMaxDataRow()+1 == ws1.getCells().getLastDataRow(13)+1) &&
                        (ws1.getCells().getMaxDataRow()+1 == ws1.getCells().getLastDataRow(14)+1) &&
                        (ws2.getCells().getMaxDataRow()+1 == ws2.getCells().getLastDataRow(1)+1) &&
                        (ws2.getCells().getMaxDataRow()+1 == ws2.getCells().getLastDataRow(3)+1) &&
                        (ws2.getCells().getMaxDataRow()+1 == ws2.getCells().getLastDataRow(5)+1) &&
                        (ws2.getCells().getMaxDataRow()+1 == ws2.getCells().getLastDataRow(7)+1) &&
                        (ws2.getCells().getMaxDataRow()+1 == ws2.getCells().getLastDataRow(11)+1) &&
                        (ws2.getCells().getMaxDataRow()+1 == ws2.getCells().getLastDataRow(17)+1) &&
                        (ws3.getCells().getMaxDataRow()+1 == ws3.getCells().getLastDataRow(1)+1) &&
                        (ws3.getCells().getMaxDataRow()+1 == ws3.getCells().getLastDataRow(3)+1) &&
                        (ws3.getCells().getMaxDataRow()+1 == ws3.getCells().getLastDataRow(5)+1) &&
                        (ws3.getCells().getMaxDataRow()+1 == ws3.getCells().getLastDataRow(7)+1) &&
                        (ws3.getCells().getMaxDataRow()+1 == ws3.getCells().getLastDataRow(11)+1)
                        ) {
                            AtomicBoolean threadsError = new AtomicBoolean(false);
                            final String[] error = {""};

                            String valPorcion = ws1.getCells().get(1,7).getStringValue();

                            Thread sheet1 = new Thread(() -> {
                                String[] errorType = {"","","","","","",""};

                                boolean t1Error = false;

                                List<String> avisos = new ArrayList<>();
                                List<Integer> pagos = new ArrayList<>();
                                List<String> tipoSolicitud = new ArrayList<>();
                                List<String> porcion = new ArrayList<>();
                                List<String> fecha = new ArrayList<>();
                                List<String> f_ejecutado = new ArrayList<>();
                                List<String> f_cierre = new ArrayList<>();

                                DateFormat dateFormat = new SimpleDateFormat("d/MM/yyyy");
                                SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");

                                for (int i = 1; i < (ws1.getCells().getLastDataRow(0)+1); i++) {
                                    if (ws1.getCells().get(i,14).getValue() != null && ws1.getCells().get(i,14).getValue() != "") {
                                        avisos.add(ws1.getCells().get(i,14).getStringValue());
                                    } else {
                                        t1Error = true;
                                        errorType[0] = "\n• numero de aviso";
                                    }

                                    try {
                                        pagos.add(ws1.getCells().get(i,1).getIntValue());
                                    } catch (Exception e) {
                                        t1Error = true;
                                        errorType[1] = "\n• pagos";
                                    }

                                    if (ws1.getCells().get(i,3).getValue() != null && ws1.getCells().get(i,3).getValue() != "" && (ws1.getCells().get(i,3).getValue().equals("SUSPENSION") || ws1.getCells().get(i,3).getValue().equals("TAPONAMIENTO"))) {
                                        tipoSolicitud.add(ws1.getCells().get(i,3).getStringValue());
                                    } else {
                                        t1Error = true;
                                        errorType[2] = "\n• tipo de solicitud";
                                    }

                                    if (ws1.getCells().get(i,7).getValue() != null && ws1.getCells().get(i,7).getValue() != "" && ws1.getCells().get(i,7).getValue() == valPorcion) {
                                        porcion.add(ws1.getCells().get(i,7).getStringValue());
                                    } else {
                                        t1Error = true;
                                        errorType[3] = "\n• porcion";
                                    }

                                    try {
                                        String validFecha = ws1.getCells().get(i,5).getStringValue();
                                        Date date = dateFormat.parse(validFecha);
                                        validFecha = simpleDateFormat.format(date);
                                        if (validFecha.length() == 10) {
                                            fecha.add(validFecha);
                                        } else {
                                            t1Error = true;
                                            errorType[5] = "\n• fecha de programación";
                                        }
                                    } catch (Exception e) {
                                        t1Error = true;
                                        errorType[4] = "\n• fecha de programación";
                                    }

                                    try {
                                        String validFecha = ws1.getCells().get(i,12).getStringValue();
                                        Date date = dateFormat.parse(validFecha);
                                        validFecha = simpleDateFormat.format(date);
                                        if (validFecha.length() == 10) {
                                            f_ejecutado.add(validFecha);
                                        } else {
                                            t1Error = true;
                                            errorType[5] = "\n• fecha de ejecutado";
                                        }

                                    } catch (Exception e) {
                                        t1Error = true;
                                        errorType[5] = "\n• fecha de ejecutado";
                                    }

                                    try {
                                        String validFecha = ws1.getCells().get(i,13).getStringValue();
                                        Date date = dateFormat.parse(validFecha);
                                        validFecha = simpleDateFormat.format(date);
                                        if (validFecha.length() == 10) {
                                            f_cierre.add(validFecha);
                                        } else {
                                            t1Error = true;
                                            errorType[5] = "\n• fecha de cierre";
                                        }
                                    } catch (Exception e) {
                                        t1Error = true;
                                        errorType[6] = "\n• fecha de cierre";
                                    }

                                }

                                if (t1Error != true) {
                                    dataWS1.add(avisos);
                                    dataWS1.add(pagos);
                                    dataWS1.add(tipoSolicitud);
                                    dataWS1.add(porcion);
                                    dataWS1.add(fecha);
                                    dataWS1.add(f_ejecutado);
                                    dataWS1.add(f_cierre);
                                } else {
                                    error[0] += "\n➤ IMPRESIÓN" + errorType[0] + errorType[1] + errorType[2] + errorType[3] + errorType[4] + errorType[5] + errorType[6] + "\n";
                                    threadsError.set(true);
                                }
                            });
                            Thread sheet2 = new Thread(() -> {
                                String[] errorType = {"","","","","",""};

                                boolean t2Error = false;

                                List<String> avisos = new ArrayList<>();
                                List<Integer> pagos = new ArrayList<>();
                                List<String> tipoSolicitud = new ArrayList<>();
                                List<String> porcion = new ArrayList<>();
                                List<Integer> resultado = new ArrayList<>();
                                List<String> fecha = new ArrayList<>();

                                DateFormat dateFormat = new SimpleDateFormat("d/MM/yyyy");
                                SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");

                                for (int i = 1; i < (ws2.getCells().getLastDataRow(0)+1); i++) {
                                    if (ws2.getCells().get(i,17).getValue() != null && ws2.getCells().get(i,17).getValue() != "") {
                                        avisos.add(ws2.getCells().get(i,17).getStringValue());
                                    } else {
                                        t2Error = true;
                                        errorType[0] = "\n• numero de aviso";
                                    }

                                    try {
                                        String value = ws2.getCells().get(i,1).getStringValue();
                                        if (value.equals("#N/A") || value.contains("$-")) {
                                            pagos.add(0);
                                        } else {
                                            pagos.add(ws2.getCells().get(i,1).getIntValue());
                                        }

                                    } catch (Exception e) {
                                        System.out.println(e);
                                        t2Error = true;
                                        errorType[1] = "\n• pagos";
                                    }

                                    if (ws2.getCells().get(i,3).getValue() != null && ws2.getCells().get(i,3).getValue() != "" && (ws2.getCells().get(i,3).getValue().equals("SUSPENSION") || ws2.getCells().get(i,3).getValue().equals("TAPONAMIENTO"))) {
                                        tipoSolicitud.add(ws2.getCells().get(i,3).getStringValue());
                                    } else {
                                        t2Error = true;
                                        errorType[2] = "\n• tipo de solicitud";
                                    }

                                    if (ws2.getCells().get(i,7).getValue() != null && ws2.getCells().get(i,7).getValue() != "" && ws2.getCells().get(i,7).getValue() == valPorcion) {
                                        porcion.add(ws2.getCells().get(i,7).getStringValue());
                                    } else {
                                        t2Error = true;
                                        errorType[3] = "\n• porcion";
                                    }

                                    try {
                                        String validFecha = ws2.getCells().get(i,5).getStringValue();
                                        Date date = dateFormat.parse(validFecha);
                                        validFecha = simpleDateFormat.format(date);
                                        fecha.add(validFecha);
                                    } catch (Exception e) {
                                        t2Error = true;
                                        errorType[4] = "\n• fecha de programación";
                                    }

                                    try {
                                        resultado.add(ws2.getCells().get(i,11).getIntValue());
                                    } catch (Exception e) {
                                        t2Error = true;
                                        errorType[1] = "\n• resultado";
                                    }

                                }

                                if (t2Error != true) {
                                    dataWS2.add(avisos);
                                    dataWS2.add(pagos);
                                    dataWS2.add(tipoSolicitud);
                                    dataWS2.add(porcion);
                                    dataWS2.add(fecha);
                                    dataWS2.add(resultado);
                                } else {
                                    error[0] += "\n➤ SyT" + errorType[0] + errorType[1] + errorType[2] + errorType[3] + errorType[4] + errorType[5] + "\n";
                                    threadsError.set(true);
                                }
                            });
                            Thread sheet3 = new Thread(() -> {
                                String[] errorType = {"","","","",""};

                                boolean t3Error = false;

                                List<String> avisos = new ArrayList<>();
                                List<Integer> pagos = new ArrayList<>();
                                List<String> tipoSolicitud = new ArrayList<>();
                                List<String> porcion = new ArrayList<>();
                                List<String> fecha = new ArrayList<>();

                                DateFormat dateFormat = new SimpleDateFormat("d/MM/yyyy");
                                SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");

                                for (int i = 1; i < (ws3.getCells().getLastDataRow(0)+1); i++) {
                                    if (ws3.getCells().get(i,11).getValue() != null && ws3.getCells().get(i,11).getValue() != "") {
                                        avisos.add(ws3.getCells().get(i,11).getStringValue());
                                    } else {
                                        t3Error = true;
                                        errorType[0] = "\n• numero de aviso";
                                    }

                                    try {
                                        String value = ws3.getCells().get(i,1).getStringValue();
                                        if (value.equals("#N/A") || value.contains("$-")) {
                                            pagos.add(0);
                                        } else {
                                            pagos.add(ws3.getCells().get(i,1).getIntValue());
                                        }

                                    } catch (Exception e) {
                                        t3Error = true;
                                        errorType[1] = "\n• pagos";
                                    }

                                    if (ws3.getCells().get(i,3).getValue() != null && ws3.getCells().get(i,3).getValue() != "" && (ws3.getCells().get(i,3).getValue().equals("SUSPENSION") || ws3.getCells().get(i,3).getValue().equals("TAPONAMIENTO"))) {
                                        tipoSolicitud.add(ws3.getCells().get(i,3).getStringValue());
                                    } else {
                                        t3Error = true;
                                        errorType[2] = "\n• tipo de solicitud";
                                    }

                                    if (ws3.getCells().get(i,7).getValue() != null && ws3.getCells().get(i,7).getValue() != "" && ws3.getCells().get(i,7).getValue() == valPorcion) {
                                        porcion.add(ws3.getCells().get(i,7).getStringValue());
                                    } else {
                                        t3Error = true;
                                        errorType[3] = "\n• porcion";
                                    }

                                    try {
                                        String validFecha = ws3.getCells().get(i,5).getStringValue();
                                        Date date = dateFormat.parse(validFecha);
                                        validFecha = simpleDateFormat.format(date);
                                        fecha.add(validFecha);
                                    } catch (Exception e) {
                                        t3Error = true;
                                        errorType[4] = "\n• fecha de programación";
                                    }
                                }

                                if (t3Error != true) {
                                    dataWS3.add(avisos);
                                    dataWS3.add(pagos);
                                    dataWS3.add(tipoSolicitud);
                                    dataWS3.add(porcion);
                                    dataWS3.add(fecha);
                                } else {
                                    error[0] += "\n➤ EXCLUIDAS" + errorType[0] + errorType[1] + errorType[2] + errorType[3] + errorType[4] + "\n";
                                    threadsError.set(true);
                                }
                            });

                            sheet1.run();
                            sheet2.run();
                            sheet3.run();

                            try {
                                sheet1.join();
                                sheet2.join();
                                sheet3.join();
                            } catch (InterruptedException e) {
                                throw new RuntimeException(e);
                            }

                            if (threadsError.get() != true) {
                                File impresionCSV = new File("files\\data_Impresion.csv");
                                File sytCSV = new File("files\\data_SyT.csv");
                                File excluidasCSV = new File("files\\data_Excluidas.csv");

                                Thread firstCSV = new Thread(() -> {
                                    try {
                                        PrintWriter write = new PrintWriter(impresionCSV);

                                        for (int j = 0; j < dataWS1.get(0).size(); j++) {
                                            for (int i = 0; i < dataWS1.size(); i++) {
                                                write.print(dataWS1.get(i).get(j));
                                                if (i < (dataWS1.size()-1)) {
                                                    write.print(",");
                                                } else if (i == (dataWS1.size()-1)) {
                                                    write.print("\n");
                                                }
                                            }
                                        }
                                        write.close();

                                    } catch (Exception e) {
                                        System.out.println(e);
                                    }

                                });
                                Thread secondCSV = new Thread(() -> {
                                    try {
                                        PrintWriter write = new PrintWriter(sytCSV);

                                        for (int j = 0; j < dataWS2.get(0).size(); j++) {
                                            for (int i = 0; i < dataWS2.size(); i++) {
                                                write.print(dataWS2.get(i).get(j));
                                                if (i < (dataWS2.size()-1)) {
                                                    write.print(",");
                                                } else if (i == (dataWS2.size()-1)) {
                                                    write.print("," + dataWS1.get(5).get(0) + "," + dataWS1.get(6).get(0) + "\n");
                                                }
                                            }
                                        }
                                        write.close();

                                    } catch (Exception e) {
                                        System.out.println(e);
                                    }
                                });
                                Thread thirdCSV = new Thread(() -> {
                                    try {
                                        PrintWriter write = new PrintWriter(excluidasCSV);

                                        for (int j = 0; j < dataWS3.get(0).size(); j++) {
                                            for (int i = 0; i < dataWS3.size(); i++) {
                                                write.print(dataWS3.get(i).get(j));
                                                if (i < (dataWS3.size()-1)) {
                                                    write.print(",");
                                                } else if (i == (dataWS3.size()-1)) {
                                                    write.print("," + dataWS1.get(5).get(0) + "," + dataWS1.get(6).get(0) + "\n");
                                                }
                                            }
                                        }
                                        write.close();

                                    } catch (Exception e) {
                                        System.out.println(e);
                                    }
                                });

                                firstCSV.start();
                                secondCSV.start();
                                thirdCSV.start();

                                try {
                                    firstCSV.join();
                                    secondCSV.join();
                                    thirdCSV.join();
                                } catch (InterruptedException e) {
                                    throw new RuntimeException(e);
                                }

                                try {
                                    File uploadFiles = new File("tools\\shell\\files.txt");
                                    PrintWriter write = new PrintWriter(uploadFiles);
                                    write.println(".mode csv");
                                    write.println(".open '" + new File("tools\\db\\database.db").getAbsolutePath() + "'");
                                    write.println(".import '" + impresionCSV.getAbsolutePath() + "' IMPRESION");
                                    write.println(".import '" + sytCSV.getAbsolutePath() + "' SyT");
                                    write.println(".import '" + excluidasCSV.getAbsolutePath() + "' EXCLUIDAS");
                                    write.println(".shell del '" + impresionCSV.getAbsolutePath() + "'");
                                    write.println(".shell del '" + sytCSV.getAbsolutePath() + "'");
                                    write.println(".shell del '" + excluidasCSV.getAbsolutePath() + "'");
                                    write.println(".exit");
                                    write.close();

                                    Process p = Runtime.getRuntime().exec("cmd /c cd " + new File("tools\\shell").getPath() + " && upload.cmd");
                                    p.getErrorStream().close();
                                    p.waitFor();

                                    String deleteDate = dataWS1.get(4).get(0).toString();
                                    try {
                                        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
                                        Calendar c = Calendar.getInstance();
                                        c.setTime(dateFormat.parse(deleteDate));
                                        c.add(Calendar.YEAR, -1);
                                        deleteDate = dateFormat.format(c.getTime());
                                    } catch (Exception e) {
                                        System.out.println(e);
                                    }

                                    conexion database = new conexion();
                                    Connection con = database.conectarSQL();

                                    String[] tables = {"IMPRESION", "SyT", "EXCLUIDAS"};
                                    for (int i = 0; i < tables.length; i++) {
                                        PreparedStatement ps = con.prepareStatement("DELETE FROM "+tables[i]+" WHERE (fecha < '"+deleteDate+"');");
                                        ps.executeUpdate();
                                    }
                                } catch (Exception e) {
                                    System.out.println(e);
                                }
                            } else {
                                typeError = error[0];
                                isError = true;
                                codeError = 5;
                            }
                        } else {
                            isError = true;
                            codeError = 4;
                        }
                    } else {
                        codeError = 3;
                        isError = true;
                        if (ws1.getCells().get(0,1).getType() != 5 || ws1.getCells().get(0,3).getType() != 5 || ws1.getCells().get(0,5).getType() != 5 || ws1.getCells().get(0,7).getType() != 5 || ws1.getCells().get(0,12).getType() != 5 && ws1.getCells().get(0,13).getType() != 5) {
                            typeError += "\n➤ IMPRESIÓN";
                        }

                        if (ws2.getCells().get(0,1).getType() != 5 || ws2.getCells().get(0,3).getType() != 5 || ws2.getCells().get(0,5).getType() != 5 || ws2.getCells().get(0,7).getType() != 5 || ws2.getCells().get(0,11).getType() != 5 || ws2.getCells().get(0,17).getType() != 5) {
                            typeError += "\n➤ SyT";
                        }

                        if ((ws3.getCells().get(0,1).getType() != 5 || ws3.getCells().get(0,3).getType() != 5 || ws3.getCells().get(0,5).getType() != 5 || ws3.getCells().get(0,7).getType() != 5 || ws3.getCells().get(0,11).getType() != 5)) {
                            typeError += "\n➤ EXCLUIDAS";
                        }
                    }
                } else {
                    codeError = 3;
                    isError = true;
                    if (ws1.getCells().getMaxDataColumn()+1 != 17) {
                        typeError += "\n➤ IMPRESIÓN";
                    }
                    if (ws2.getCells().getMaxDataColumn()+1 != 21) {
                        typeError += "\n➤ SyT";
                    }
                    if (ws3.getCells().getMaxDataColumn()+1 != 13) {
                        typeError += "\n➤ EXCLUIDAS";
                    }
                }
            } else {
                codeError = 2;
                isError = true;
                typeError = "\n\n1) IMPRESION\n2) SyT \n3) EXCLUIDAS";
            }
        } else {
            codeError = 1;
            isError = true;
        }
    }

    public void upload (File file, Stage initStage, TextField tf) {
        try {
            Workbook wbXLSX = new Workbook(file.getAbsolutePath());
            new Thread(() -> {validXLSX(wbXLSX);}).run();
        } catch (Exception e) {
            System.out.println(e);
        }

        Platform.runLater(new Runnable() {
            @Override
            public void run() {
                new loading(initStage);
                Stage primaryStage = new Stage();
                new app().start(primaryStage);

                if (isError != true) {
                    tf.setText(null);
                    Alert alert = new Alert(Alert.AlertType.INFORMATION);
                    alert.setHeaderText(null);
                    alert.setTitle("Success");
                    alert.setContentText(file.getName()+" SUBIDO CORRECTAMENTE.");
                    alert.showAndWait();
                } else {
                    if (codeError == 1) {
                        Alert alert = new Alert(Alert.AlertType.WARNING);
                        alert.setHeaderText(null);
                        alert.setTitle("Error");
                        alert.setContentText("ARCHIVO INVALIDO.");
                        alert.showAndWait();
                    } else if (codeError == 2) {
                        Alert alert = new Alert(Alert.AlertType.WARNING);
                        alert.setHeaderText(null);
                        alert.setTitle("Error");
                        alert.setContentText("VERIFIQUE QUE SEA UN ACTA VALIDA CON LAS HOJAS CORRESPONDIENTES Y EL SIGUIENTE ORDEN: " + typeError);
                        alert.showAndWait();
                    } else if (codeError == 3) {
                        Alert alert = new Alert(Alert.AlertType.WARNING);
                        alert.setHeaderText(null);
                        alert.setTitle("Error");
                        alert.setContentText("VERIFIQUE LA ESTRUCTURA DE LA(s) HOJA(s):\n" + typeError);
                        alert.showAndWait();
                    } else if (codeError == 4) {
                        Alert alert = new Alert(Alert.AlertType.WARNING);
                        alert.setHeaderText(null);
                        alert.setTitle("Error");
                        alert.setContentText("VERIFIQUE LOS DATOS DEL ARCHIVO.");
                        alert.showAndWait();
                    } else if (codeError == 5) {
                        Alert alert = new Alert(Alert.AlertType.WARNING);
                        alert.setHeaderText(null);
                        alert.setTitle("Error");
                        alert.setContentText("VERIFIQUE LOS SIGUIENTES CAMPOS:\n" + typeError);
                        alert.showAndWait();
                    }
                }
            }
        });
    }
}