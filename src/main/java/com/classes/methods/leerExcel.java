package com.classes.methods;

import com.classes.connection.conexion;
import javafx.stage.Stage;

import javax.swing.*;
import java.io.File;
import java.io.PrintWriter;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

public class leerExcel {
    public void leerExcel(File file, Stage primaryStage) {
            loading load = new loading();
            primaryStage.close();
            load.loader();
        try {
            Workbook wbXLSX = new Workbook(PATH); //NUEVO LIBRO EXCEL
            Worksheet ws = wbXLSX.getWorksheets().get(0); //HOJA EXCEL, PRIMERA HOJA
            //VALIDAR ESTRUCTURA
            int valCOLUMN = ws.getCells().getMaxDataColumn(); //RECUENTO DE COLUMNA
            //SI TIENE 21 COLUMNAS HACER ESTO
            if ((valCOLUMN+1) == 21) {
                int valROW1 = ws.getCells().getLastDataRow(0); //RECUENTO DE COLUMNAS
                int valROW2 = ws.getCells().getMaxDataRow();
                if (valROW1 == valROW2) {
                    File fileDATA = new File("files\\Importe.csv"); //CREAR UN NUEVO ARCHIVO EN LA CARPETA files CON EL NOMBRE DE Importe DE TIPO csv
                    wbXLSX.save("" + fileDATA); //GUARDAR LOS DATOS DEL LIBRO EN EL ARCHIVO csv
                    String rutaCSV = "" + fileDATA; //GUARDAR RUTA EN UNA VARIABLE
                    // 2. LEE LOS DATOS DEL ARCHIVO Y LOS GUARDA EN UNA LISTA
                    List<LECTURAS> DATA; //LISTA CON MODELO DE LECTURAS LLAMADA DATA
                    DATA = new ArrayList<>(); //NUEVA LISTA DE DATOS DONDE SE GUARDARAN LOS DATOS DEL ARCHIVO
                    CsvReader readLECTURAS = new CsvReader(rutaCSV);
                    readLECTURAS.readHeaders();
                    //CICLO QUE LEE CADA DATO DEL ARCHIVO Y LOS ALMACENA EN LA LISTA
                    while (readLECTURAS.readRecord()) {
                        String codigo_porcion = readLECTURAS.get(0);
                        String uni_lectura = readLECTURAS.get(1);
                        String doc_lectura = readLECTURAS.get(2);
                        String cuenta_contrato = readLECTURAS.get(3);
                        String medidor = readLECTURAS.get(4);
                        String lectura_ant = readLECTURAS.get(5);
                        String lectura_act = readLECTURAS.get(6);
                        String anomalia_1 = readLECTURAS.get(7);
                        String anomalia_2 = readLECTURAS.get(8);
                        String codigo_operario = readLECTURAS.get(9);
                        String vigencia = readLECTURAS.get(10);
                        //CONVERTIR LOS DATOS RECIBIDOS DE fecha CON FORMATO yyyy/MM/dd HH:mm PARA MEJORAR LA FILTRACION
                        String fecha = readLECTURAS.get(11);
                        Calendar gregorianCalendar = new GregorianCalendar();
                        DateFormat dateFormat = new SimpleDateFormat("d/MM/yyyy HH:mm");
                        Date date = dateFormat.parse(fecha);
                        gregorianCalendar.setTime(date);
                        Locale locale = new Locale("es", "EC");
                        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm", locale);
                        fecha = simpleDateFormat.format(date);
                        //
                        String orden_lectura = readLECTURAS.get(12);
                        String leido = readLECTURAS.get(13);
                        String calle = readLECTURAS.get(14);
                        String edificio = readLECTURAS.get(15);
                        String suplemento_casa = readLECTURAS.get(16);
                        String interloc_comercial = readLECTURAS.get(17);
                        String apellido = readLECTURAS.get(18);
                        String nombre = readLECTURAS.get(19);
                        String clase_instalacion = readLECTURAS.get(20);

                        //SI EL DATO TIENE COMA, ELIMINARLA
                        codigo_porcion = codigo_porcion.replaceAll(",", "");
                        uni_lectura = uni_lectura.replaceAll(",", "");
                        doc_lectura = doc_lectura.replaceAll(",", "");
                        cuenta_contrato = cuenta_contrato.replaceAll(",", "");
                        medidor = medidor.replaceAll(",", "");
                        lectura_ant = lectura_ant.replaceAll(",", "");
                        lectura_act = lectura_act.replaceAll(",", "");
                        anomalia_1 = anomalia_1.replaceAll(",", "");
                        anomalia_2 = anomalia_2.replaceAll(",", "");
                        codigo_operario = codigo_operario.replaceAll(",", "");
                        vigencia = vigencia.replaceAll(",", "");
                        fecha = fecha.replaceAll(",", "");
                        orden_lectura = orden_lectura.replaceAll(",", "");
                        leido = leido.replaceAll(",", "");
                        calle = calle.replaceAll(",", "");
                        edificio = edificio.replaceAll(",", "");
                        suplemento_casa = suplemento_casa.replaceAll(",", "");
                        interloc_comercial = interloc_comercial.replaceAll(",", "");
                        apellido = apellido.replaceAll(",", "");
                        nombre = nombre.replaceAll(",", "");

                        if (codigo_porcion == "" || uni_lectura == "" || codigo_operario == "" || vigencia == "") {
                            dialog.dispose();
                            JOptionPane.showMessageDialog(null, "ERROR: VERIFIQUE QUE LOS CAMPOS CODIGO PORCION, UNI LECTURA, CODIGO OPERARIO O VIGENCIA NO SE ENCUENTREN VACIOS", "", JOptionPane.INFORMATION_MESSAGE);
                            return;
                        }

                        //SI EL DATO TIENE COMILLAS, ELIMINARLAS
                        codigo_porcion = codigo_porcion.replaceAll("\"", "");
                        uni_lectura = uni_lectura.replaceAll("\"", "");
                        doc_lectura = doc_lectura.replaceAll("\"", "");
                        cuenta_contrato = cuenta_contrato.replaceAll("\"", "");
                        medidor = medidor.replaceAll("\"", "");
                        lectura_ant = lectura_ant.replaceAll("\"", "");
                        lectura_act = lectura_act.replaceAll("\"", "");
                        anomalia_1 = anomalia_1.replaceAll("\"", "");
                        anomalia_2 = anomalia_2.replaceAll("\"", "");
                        codigo_operario = codigo_operario.replaceAll("\"", "");
                        vigencia = vigencia.replaceAll("\"", "");
                        fecha = fecha.replaceAll("\"", "");
                        orden_lectura = orden_lectura.replaceAll("\"", "");
                        leido = leido.replaceAll("\"", "");
                        calle = calle.replaceAll("\"", "");
                        edificio = edificio.replaceAll("\"", "");
                        suplemento_casa = suplemento_casa.replaceAll("\"", "");
                        interloc_comercial = interloc_comercial.replaceAll("\"", "");
                        apellido = apellido.replaceAll("\"", "");
                        nombre = nombre.replaceAll("\"", "");

                        if (codigo_porcion.charAt(0) == 'W' || codigo_porcion.charAt(0) == 'X' || codigo_porcion.charAt(0) == 'Z') {
                            if (vigencia.charAt(5) == '0') {
                                codigo_porcion += "-1";
                            } else {
                                codigo_porcion += "-2";
                                StringBuilder nuevaVIGENCIA = new StringBuilder(vigencia);
                                nuevaVIGENCIA.setCharAt(5,'0');
                                vigencia = nuevaVIGENCIA.toString();
                            }
                        }
                        DATA.add(new LECTURAS(codigo_porcion, uni_lectura, doc_lectura, cuenta_contrato, medidor, lectura_ant, lectura_act, anomalia_1, anomalia_2, codigo_operario, vigencia, fecha, orden_lectura, leido, calle, edificio, suplemento_casa, interloc_comercial, apellido, nombre, clase_instalacion));
                    }
                    readLECTURAS.close();

                    //EXTRAER DATOS REPETIDOS DEL ARCHIVO
                    Set<LECTURAS> repetidos; //SET CON MODELO LECTURAS
                    repetidos = new HashSet<>(); //HASHSET PARA SACAR LOS REPETIDOS
                    List<LECTURAS> repetidosFinal; //LISTA CON MODELO LECTURAS
                    repetidosFinal = DATA.stream().filter(lectura -> !repetidos.add(lectura)).collect(Collectors.toList()); //GUARDAR DATOS REPETIDOS EN LA LISTA

                    boolean fileOPEN = false;
                    String name = jtxtPATH.getText();
                    name = name.replaceAll(" ", "_");
                    File fileNAME = new File(name);

                    //SI HAY REPETIDOS EXPORTARLOS EN UN EXCEL
                    if (repetidosFinal.size() != 0) {
                        File fileREPLY = new File("files\\Repetidos.csv"); //ARCHIVO PARA RETORNAR REPETIDOS EN UN ARCHIVO csv
                        PrintWriter write = new PrintWriter(fileREPLY); //PARA ESCRIBIR LOS DATOS REPETIDOS EN EL NUEVO ARCHIVO

                        String estructura = "CODIGO_PORCION,UNI_LECTURA,DOC_LECTURA,CUENTA_CONTRATO,MEDIDOR,LEC_ANTERIOR,LEC_ACTUAL,ANOMALIA_1,ANOMALIA_2,CODIGO_OPERARIO,VIGENCIA,FECHA,ORDEN LECTURA,LEIDO,CALLE,EDIFICIO,SUPLEMENTO_CASA,INTERLOC_COM,APELLIDO,NOMBRE,CLASE_INSTALA";
                        write.println(estructura);

                        for (Modelo.LECTURAS LECTURAS : repetidosFinal) {
                            write.print(LECTURAS.getCodigo_porcion() + ",");
                            write.print(LECTURAS.getUni_lectura() + ",");
                            write.print(LECTURAS.getDoc_lectura() + ",");
                            write.print(LECTURAS.getCuenta_contrato() + ",");
                            write.print(LECTURAS.getMedidor() + ",");
                            write.print(LECTURAS.getLectura_ant() + ",");
                            write.print(LECTURAS.getLectura_act() + ",");
                            write.print(LECTURAS.getAnomalia_1() + ",");
                            write.print(LECTURAS.getAnomalia_2() + ",");
                            write.print(LECTURAS.getCodigo_operario() + ",");
                            write.print(LECTURAS.getVigencia() + ",");
                            write.print(LECTURAS.getFecha() + ",");
                            write.print(LECTURAS.getOrden_lectura() + ",");
                            write.print(LECTURAS.getLeido() + ",");
                            write.print(LECTURAS.getCalle() + ",");
                            write.print(LECTURAS.getEdificio() + ",");
                            write.print(LECTURAS.getSuplemento_casa() + ",");
                            write.print(LECTURAS.getInterloc_comercial() + ",");
                            write.print(LECTURAS.getApellido() + ",");
                            write.print(LECTURAS.getNombre() + ",");
                            write.print(LECTURAS.getClase_instalacion());
                            write.println();
                        }
                        write.close();
                        //TRATAR DE CONVERTIR EL ARCHIVO.CSV CON DATOS REPETIDOS EN UN ARCHIVO.XLSX
                        try {
                            Workbook wbCSV = new Workbook("files\\Repetidos.csv"); //NUEVO LIBRO DEL ARCHIVO Repetidos
                            wbCSV.save("files\\REPETIDOS_" + fileNAME.getName(), SaveFormat.XLSX); //GUARDAR DATOS REPETIDOS EN UN ARCHIVO EXCEL
                        } catch (Exception e) {
                            fileOPEN = true;
                            dialog.dispose(); //CERRAR LOADING
                            JOptionPane.showMessageDialog(null, "ERROR: EL ARCHIVO NO PUEDE SER IMPORTADO PORQUE UN ARCHIVO RELACIONADO A LOS REGISTROS REPETIDOS SE ENCUENTRA ABIERTO", "", JOptionPane.INFORMATION_MESSAGE);
                        }
                        fileREPLY.delete(); //ELIMINAR ARCHIVO DE Repetidos.csv
                    }
                    fileDATA.delete(); //ELIMINAR ARCHIVO DE Importe.csv

                    if (fileOPEN != true) {
                        DATA = DATA.stream().distinct().collect(Collectors.toList()); //GUARDAR DATOS COMPLETOS SIN REPETIDOS
                        File RutaDATA = new File("files\\Datos.csv"); //ARCHIVO CON LOS DATOS COMPLETOS EN FORMATO csv
                        PrintWriter writeDATA = new PrintWriter(RutaDATA); //PARA ESCRIBIR LOS DATOS COMPLETOS EN EL NUEVO ARCHIVO

                        for (Modelo.LECTURAS LECTURAS : DATA) {
                            writeDATA.print(LECTURAS.getCodigo_porcion() + ",");
                            writeDATA.print(LECTURAS.getUni_lectura() + ",");
                            writeDATA.print(LECTURAS.getDoc_lectura() + ",");
                            writeDATA.print(LECTURAS.getCuenta_contrato() + ",");
                            writeDATA.print(LECTURAS.getMedidor() + ",");
                            writeDATA.print(LECTURAS.getLectura_ant() + ",");
                            writeDATA.print(LECTURAS.getLectura_act() + ",");
                            writeDATA.print(LECTURAS.getAnomalia_1() + ",");
                            writeDATA.print(LECTURAS.getAnomalia_2() + ",");
                            writeDATA.print(LECTURAS.getCodigo_operario() + ",");
                            writeDATA.print(LECTURAS.getVigencia() + ",");
                            writeDATA.print(LECTURAS.getFecha() + ",");
                            writeDATA.print(LECTURAS.getOrden_lectura() + ",");
                            writeDATA.print(LECTURAS.getLeido() + ",");
                            writeDATA.print(LECTURAS.getCalle() + ",");
                            writeDATA.print(LECTURAS.getEdificio() + ",");
                            writeDATA.print(LECTURAS.getSuplemento_casa() + ",");
                            writeDATA.print(LECTURAS.getInterloc_comercial() + ",");
                            writeDATA.print(LECTURAS.getApellido() + ",");
                            writeDATA.print(LECTURAS.getNombre() + ",");
                            writeDATA.print(LECTURAS.getClase_instalacion());
                            writeDATA.println();
                        }
                        writeDATA.close();

                        // 3. IMPORTAR LISTA DE DATOS A LA BASE DE DATOS
                        //CREAR ARCHIVO DE COMANDOS CON LAS RUTAS DE LA BASE DE DATOS Y EL ARCHIVO
                        File RutaCARPETA = new File("lib\\sqlite-tools");
                        File RutaCOMANDOS = new File("lib\\sqlite-tools\\comandos.txt");
                        PrintWriter writeCOMANDOS = new PrintWriter(RutaCOMANDOS); //PARA ESCRIBIR EL COMANDO CON LA RUTA DE LOS DATOS

                        //COMANDO (script)
                        writeCOMANDOS.println(".mode csv");
                        writeCOMANDOS.println(".open '" + pathDB.getAbsolutePath() + "'");
                        writeCOMANDOS.println(".import '" + RutaDATA.getAbsolutePath() + "' LECTURAS");
                        writeCOMANDOS.println(".shell del '" + RutaDATA.getAbsolutePath() + "'");
                        writeCOMANDOS.close();

                        //LINEA DE COMANDOS EJECUTANDO EL COMANDO (script)
                        Runtime.getRuntime().exec("cmd /c cd " + RutaCARPETA.getAbsolutePath() + " && script.cmd");
                        Thread.sleep(2*1000);

                        //RESETEAR LOS DATOS PARA FILTRAR Y GENERAR INFORME E INICIAR METODO INIT
                        jpSCROLL_CODPOR.removeAll();
                        puMENU_CODPOR.removeAll();
                        jpSCROLL_RUTAS.removeAll();
                        puMENU_RUTAS.removeAll();
                        jpSCROLL_CODOPE.removeAll();
                        puMENU_CODOPE.removeAll();
                        jpSCROLL_VIG.removeAll();
                        puMENU_VIG.removeAll();
                        new Thread (()-> INIT()).run();

                        JOptionPane.showMessageDialog(null, "SE IMPORTO CORRECTAMENTE " + DATA.size() + " REGISTROS DE " + fileNAME.getName(), "", JOptionPane.INFORMATION_MESSAGE);
                        if (repetidosFinal.size() != 0) {
                            JOptionPane.showMessageDialog(null, "SE ENCONTRARON " + repetidosFinal.size() + " REGISTROS REPETIDOS EN " + fileNAME.getName(), "", JOptionPane.INFORMATION_MESSAGE);
                            File rutaARCHIVOS = new File("files");
                            Runtime.getRuntime().exec("cmd /c start " + rutaARCHIVOS.getAbsolutePath() + "\\REPETIDOS_" + fileNAME.getName() + " && exit");
                        }
                    }
                } else {
                    dialog.dispose(); //CERRAR LOADING
                    JOptionPane.showMessageDialog(null, "ERROR: VERIFIQUE LOS DATOS DEL ARCHIVO", "",JOptionPane.INFORMATION_MESSAGE); //MENSAJE DE ERROR POR DATOS MAL ESCRITOS EN ALGUNAS COLUMNAS
                }
            } else {
                dialog.dispose(); //CERRAR LOADING
                JOptionPane.showMessageDialog(null, "ERROR: VERIFIQUE LA ESTRUCTURA DEL ARCHIVO", "",JOptionPane.INFORMATION_MESSAGE); //MENSAJE DE ERROR POR LA ESTRUCTURA DEL ARCHIVO
            }
        } catch (Exception e) {
            dialog.dispose(); //CERRAR LOADING
            File file = new File("files\\Importe.csv");
            file.delete();
            JOptionPane.showMessageDialog(null, "ERROR: VERIFIQUE LAS FECHAS DEL ARCHIVO", "",JOptionPane.INFORMATION_MESSAGE); //MENSAJE DE ERROR POR DATOS MAL ESCRITOS EN ALGUNAS COLUMNAS
            loading close = new loading();
            close.closeLoader();
        }
    }
}
