/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package prueba2;

import java.io.FileOutputStream;
import java.io.IOException;
import java.net.InetAddress;
import java.net.UnknownHostException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;
  
import com.google.gson.JsonArray;
import com.google.gson.JsonObject;
//import common.PrintLog;
//import constants.Iconstants;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.net.InetAddress;
import java.net.UnknownHostException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;
//import ticofoniaMessenger.Mandrill;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 *
 * @author gil
 */
public class GenerateExcel { 

    private transient final String datePattern = "yyyy-MM-dd";
    private transient SimpleDateFormat dateFormat = new SimpleDateFormat(datePattern);
    private transient boolean production;

    public GenerateExcel() {
        try {

            String IP = InetAddress.getLocalHost().getHostAddress();
            String Name = InetAddress.getLocalHost().getHostName();
            // PrintLog.Print("IP:" + IP + " Name:" + Name);
            int index = Name.indexOf("moovin");
            int indexDeveloper = Name.indexOf("developer");
            if (indexDeveloper != -1) {
                production = false;
            } else if (index != -1) {
                production = true;
                //key = keyDeveloper;
            } else {
                //key = keyProduction;
                production = false;
            }

            /*int indexDeveloper = Name.indexOf("developer");
                if (indexDeveloper != -1) {
                    key = keyDeveloper;

                    //key = keyDeveloper;
                } else {
                    //key = keyProduction;
                    key = keyProduction;
                }*/
            //ably = new AblyRealtime(key);
            // PrintLog.Print("URL: " + production);
        } catch (UnknownHostException ex) {
            production = true;
        }
    }

    public String saveExcel(String[] headers, ArrayList<ArrayList> dataInfo, String nameReport, String dateStart, String dateEnd) {
        // reference: https://hashblogeando.wordpress.com/2016/02/05/creando-archivos-excel-en-formato-xlsx/ 
//        String fileURL = "/home/developer/Desktop/reports/";
        String fileURL = "/Users/gil/Desktop/reports/";              // se debe de cambiar.
        System.out.println("entro a crear el excel");
        int positionHeader = 0;
        int initialPosititionData = positionHeader + 1;
        
        try {
            // --------------------- instancia de excel ---------------------
            // Creacion del libro de excel.
            HSSFWorkbook workbook = new HSSFWorkbook();
            
            // creamos la hoja donde se pondra datos
            HSSFSheet sheet = workbook.createSheet();
            System.out.println("entro a crear el excel 0");
            // nombre de la hoja a trabajar
            workbook.setSheetName(0, "Hoja");
            System.out.println("entro a crear el excel 1");
           
            // ------------------- estilo del encabezado -------------------
            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.SEA_GREEN.getIndex());  // agregarle un color al fondo de las celdas.
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setAlignment(HorizontalAlignment.CENTER);
//            headerStyle.setFillPattern(CellStyle.class.);            // el fondo con patron solido del color indicado.
            // modificacion de la fuente.
            Font font = workbook.createFont();      
            font.setBold(true);
            headerStyle.setFont(font);
            // -------------------------------------------------------------

            System.out.println("entro a crear el excel 2");
            CellStyle style = workbook.createCellStyle();
            style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());

            // ----------------- ANADIMOS DATOS AL HEADER ----------------------
            // se crea una fila en la hoja en la posicion 0.
            HSSFRow headerRow = sheet.createRow(positionHeader); 
            // creamos el encabezado
            for (int i = 0; i < headers.length; ++i) {
                String header = headers[i];
                // creamos celda en la fila creada anterior en la posicion del for.
                HSSFCell cell = headerRow.createCell(i);
                // le aplicamos el estilo configurado
                cell.setCellStyle(headerStyle);
                cell.setCellValue(header);
            }
            // -----------------------------------------------------------------

            System.out.println("entro a crear el excel 3");
//            int numColums = headers.length;

            for (int i = 0; i < dataInfo.size(); ++i) {
                // recorrer informacion obtenida
                ArrayList<String> row = dataInfo.get(i); 
                // creamos una fila nueva para los datos anteriores.
                HSSFRow dataRow = sheet.createRow(i + initialPosititionData );
                String formula ;
                System.out.print("tamano fila: "+ row.size());
                for (int j = 0; j < row.size(); j++) {
                    // guardamos la info en la posicion (celda) dada por el for.
                    int col = i+2 ; // donde inicia los datos a escribirse
                    switch(j){
//                        case 7:         // columna g: VALIDACION PLAN
//                            formula = "BUSCARV(E:E;VALIDACION!A:B;2;FALSO)" ;
//                            dataRow.createCell(j).setCellFormula( formula );  
//                            break;
//                        case 13:         // columna L: MONTO DE PLAN
//                            formula = "BUSCARV(F:F;MONTOS!A:B;2;FALSO)" ;
//                            dataRow.createCell(j).setCellFormula( formula );  
//                            break;
                        case 14:         // columna O: MONTO COMISION
                            System.out.print("tamano comision: "+ row.get(j));
                            char L = 'L';
                            char M = 'M';
                            char N = 'N';
                            formula = "("+M+col+ "-(" +M+col+ "*"+L+col+"/100 ) )*"+N+col+ "/100"; //  =(H6-(H6*11,95%))*10% 
                            System.out.println(formula);
                            dataRow.createCell(j).setCellFormula(formula );  
                            break;
                        
                         default: 
                            dataRow.createCell(j).setCellValue(row.get(j));  
                            break; 
                    } 
                    // auto ajustador de columnas
                    sheet.autoSizeColumn(j);
                     
                    /* Lo siguiente es un ejm de como agregar una funcion a una celda:
                    --> https://www.adictosaltrabajo.com/2011/11/02/hojas-calculo-formulas-poi/
                    
                    - para adjuntar un rango (como el excel de leo):
                        - con size del array obtener el rango de datos y setear el rango en la formula (la misma de leo en el excel):
                    
                    */
//                    dataRow.createCell(j).setCellFormula(fileURL);
                    // entonces con lo anterior, podemos configurar para cada celda lo siguiente:
                     
//                    dataRow.createCell(j).setCellFormula('SUM(B2:F2)');
                }   
            }

            System.out.println("entro a crear el excel 4");

            Date date = new Date();
            String nameReportSave = nameReport + "-" + dateFormat.format(date) + ".xls";
            fileURL += nameReportSave;
            try {
                System.out.println("entro a crear el excel 5");

                // guardamos el archivo.
                FileOutputStream file = new FileOutputStream(fileURL);
                workbook.write(file);
                file.close();
                System.out.println("entro a crear el excel 6");
            } catch (IOException ex) {
                Logger.getLogger(GenerateExcel.class.getName()).log(Level.SEVERE, null, ex);
                System.out.println("Error " + ex);
            }
        } catch (Exception ex) {
            System.out.println("Error " + ex);
        }
        return fileURL;
    }
     



    // https://www.adictosaltrabajo.com/2011/11/02/hojas-calculo-formulas-poi/
    // devuelve el rango de columnas en la que actua la formula.
//    public String generaRangoFormulaEnFila(int numeroFila){
//        // la posicion(nombre) de la columna se definen por ASCI: https://elcodigoascii.com.ar//
///        final byte columnaA = 65;
//        final byte columnaB = 66;
//        final byte columnaC = 67;
//        final byte columnaD = 68;
//        final byte columnaE = 69;
//        final byte columnaF = 70;
//        
//        final char primeraColumna = (char)columnaB;
//        final char ultimaColumna = (char)columnaB + Piloto.NUMERO_VUELTAS_ENTRENAMIENTO - 1;
//        return "(" + primeraColumna + numeroFila + ":" + ultimaColumna + numeroFila + ")";
//    }

//    public void saveRoutesExcel(String[] headers, ArrayList<ArrayList> dataInfo, String nameReport, String dateStart, String dateEnd) {
//        // PrintLog.PrintError("entro a crear el excel");
//        try {
//            HSSFWorkbook workbook = new HSSFWorkbook();
//            HSSFSheet sheet = workbook.createSheet();
//            // PrintLog.PrintProduction("entro a crear el excel 0");
//            workbook.setSheetName(0, "Hoja");
//            // PrintLog.PrintProduction("entro a crear el excel 1");
//
//            CellStyle headerStyle = workbook.createCellStyle();
//            Font font = workbook.createFont();
//            font.setBold(true);
//            headerStyle.setFont(font);
//
//            // PrintLog.PrintProduction("entro a crear el excel 2");
//            CellStyle style = workbook.createCellStyle();
//            style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
//
//            HSSFRow headerRow = sheet.createRow(0);
//            for (int i = 0; i < headers.length; ++i) {
//                String header = headers[i];
//                HSSFCell cell = headerRow.createCell(i);
//                cell.setCellStyle(headerStyle);
//                cell.setCellValue(header);
//            }
//
//            // PrintLog.PrintError("entro a crear el excel 3");
//
//            for (int i = 0; i < dataInfo.size(); ++i) {
//                ArrayList<String> row = dataInfo.get(i);
//
//                HSSFRow dataRow = sheet.createRow(i + 1);
//                for (int j = 0; j < row.size(); j++) {
//                    dataRow.createCell(j).setCellValue(row.get(j));
//                }
//
//                /*dataRow.createCell(0).setCellValue(row.get(0));
//                dataRow.createCell(1).setCellValue(row.get(1));
//                dataRow.createCell(2).setCellValue(row.get(2));
//                dataRow.createCell(3).setCellValue(row.get(3));
//                dataRow.createCell(4).setCellValue(row.get(4));
//                dataRow.createCell(5).setCellValue(row.get(5));*/
//            }
//
//            // PrintLog.PrintError("entro a crear el excel 4");
//
//            Date date = new Date();
//            String nameReportSave = nameReport + "-" + dateFormat.format(date) + ".xls";
//            String url = "https://moovin.me/MoovinReports/" + nameReportSave;
//            try {
//                // PrintLog.PrintError("entro a crear el excel 5");
//
//                FileOutputStream file = new FileOutputStream("/var/www/html/MoovinReports/" + nameReportSave);
//                workbook.write(file);
//                file.close();
//                // PrintLog.PrintError("entro a crear el excel 6");
//            } catch (IOException ex) {
//                Logger.getLogger(GenerateExcel.class.getName()).log(Level.SEVERE, null, ex);
//                // PrintLog.PrintError("Error " + ex);
//            }
//           // Mandrill mandrill = new Mandrill();
//           // mandrill.sendMail("Emailing reporte genérico (uso interno)", "Reporte " + nameReportSave, "tuvidamasfacil@moovin.me", "Moovin", sendRoutesMail(), new JsonArray(), sendMergeVarsRoute(nameReportSave, url, dateStart, dateEnd));
//        } catch (Exception ex) {
//            // PrintLog.PrintError("Error " + ex);
//        }
//
//    }
//
//    public JsonArray sendMail() {
//        JsonArray listSend = new JsonArray();
//        JsonObject client = new JsonObject();
//        client.addProperty("email", "lurena@ticofonia.com");
//        //client.addProperty("email", "ejimenez@moovin.me");
//        client.addProperty("name", "Leonardo");
//        client.addProperty("type", "to");
//        listSend.add(client);
//
//        JsonObject client2 = new JsonObject();
//        client2.addProperty("email", "ibolanos@ticofonia.com");
//        //client.addProperty("email", "ejimenez@moovin.me");
//        client2.addProperty("name", "Ingrid");
//        client2.addProperty("type", "to");
//        listSend.add(client2);
//        return listSend;
//
//    }
//
//    public JsonArray sendRoutesMail() {
//        JsonArray listSend = new JsonArray();
//        if (production) {
//
//            JsonObject client = new JsonObject();
//            client.addProperty("email", "jaleman@ticofonia.com");
//            //client.addProperty("email", "ejimenez@moovin.me");
//            client.addProperty("name", "July");
//            client.addProperty("type", "to");
//            listSend.add(client);
//
//            JsonObject client2 = new JsonObject();
//            client2.addProperty("email", "Katty.pastor@moovin.me");
//            //client.addProperty("email", "ejimenez@moovin.me");
//            client2.addProperty("name", "Katty");
//            client2.addProperty("type", "to");
//            listSend.add(client2);
//        } else {
//            JsonObject client = new JsonObject();
//            client.addProperty("email", "info@moovin.me");
//            //client.addProperty("email", "ejimenez@moovin.me");
//            client.addProperty("name", "Prueba developer");
//            client.addProperty("type", "to");
//            listSend.add(client);
//        }
//        return listSend;
//
//    }
//
//    public JsonArray sendMergeVars(String name, String url, String dateStart, String dateEnd) {
//
//        // PrintLog.Print("name " + name + " url " + url);
//        JsonArray merge_vars = new JsonArray();
//        JsonArray vars = new JsonArray();
//        JsonObject info = new JsonObject();
//        info.addProperty("name", "FNAME");
//        info.addProperty("content", "Leonardo");
//        vars.add(info);
//        JsonObject info0 = new JsonObject();
//        info0.addProperty("name", "URL");
//        info0.addProperty("content", url);
//        vars.add(info0);
//        JsonObject info1 = new JsonObject();
//        info1.addProperty("name", "PERIOD1");
//        info1.addProperty("content", dateStart);
//        vars.add(info1);
//        JsonObject info3 = new JsonObject();
//        info3.addProperty("name", "PERIOD2");
//        info3.addProperty("content", dateEnd);
//        vars.add(info3);
//        JsonObject info2 = new JsonObject();
//        info2.addProperty("name", "NAMEREPORT");
//        info2.addProperty("content", name);
//        vars.add(info2);
//        JsonObject var_mail = new JsonObject();
//        var_mail.addProperty("rcpt", "lurena@ticofonia.com");
//        // PrintLog.PrintProduction("vars " + vars);
//        var_mail.add("vars", vars);
//        merge_vars.add(var_mail);
//
//        JsonArray vars1 = new JsonArray();
//        JsonObject infoperson = new JsonObject();
//        infoperson.addProperty("name", "FNAME");
//        infoperson.addProperty("content", "Ingrid");
//        vars1.add(info);
//        JsonObject infoperson0 = new JsonObject();
//        infoperson0.addProperty("name", "URL");
//        infoperson0.addProperty("content", url);
//        vars1.add(infoperson0);
//        JsonObject infoperson1 = new JsonObject();
//        infoperson1.addProperty("name", "PERIOD1");
//        infoperson1.addProperty("content", dateStart);
//        vars1.add(infoperson1);
//        JsonObject infoperson2 = new JsonObject();
//        infoperson2.addProperty("name", "PERIOD2");
//        infoperson2.addProperty("content", dateEnd);
//        vars1.add(infoperson2);
//        JsonObject infoperson3 = new JsonObject();
//        infoperson3.addProperty("name", "NAMEREPORT");
//        infoperson3.addProperty("content", name);
//        vars1.add(infoperson3);
//        JsonObject var_mail1 = new JsonObject();
//        //var_mail.addProperty("rcpt", "ejimenez@moovin.me");
//        var_mail1.addProperty("rcpt", "ibolanos@ticofonia.com");
//        // PrintLog.PrintProduction("vars " + vars1);
//        var_mail1.add("vars", vars1);
//        merge_vars.add(var_mail1);
//        return merge_vars;
//
//    }
//
//    public JsonArray sendMergeVarsRoute(String name, String url, String dateStart, String dateEnd) {
//
//        // PrintLog.Print("name " + name + " url " + url);
//        JsonArray merge_vars = new JsonArray();
//        JsonArray vars = new JsonArray();
//        JsonObject info = new JsonObject();
//        info.addProperty("name", "FNAME");
//        info.addProperty("content", "July");
//        vars.add(info);
//        JsonObject info0 = new JsonObject();
//        info0.addProperty("name", "URL");
//        info0.addProperty("content", url);
//        vars.add(info0);
//        JsonObject info1 = new JsonObject();
//        info1.addProperty("name", "PERIOD1");
//        info1.addProperty("content", dateStart);
//        vars.add(info1);
//        JsonObject info3 = new JsonObject();
//        info3.addProperty("name", "PERIOD2");
//        info3.addProperty("content", dateEnd);
//        vars.add(info3);
//        JsonObject info2 = new JsonObject();
//        info2.addProperty("name", "NAMEREPORT");
//        info2.addProperty("content", name);
//        vars.add(info2);
//        JsonObject var_mail = new JsonObject();
//        var_mail.addProperty("rcpt", "jaleman@ticofonia.com");
//        // PrintLog.Print("vars " + vars);
//        var_mail.add("vars", vars);
//        merge_vars.add(var_mail);
//
//        JsonArray vars1 = new JsonArray();
//        JsonObject infoperson = new JsonObject();
//        infoperson.addProperty("name", "FNAME");
//        infoperson.addProperty("content", "Katty");
//        vars1.add(info);
//        JsonObject infoperson0 = new JsonObject();
//        infoperson0.addProperty("name", "URL");
//        infoperson0.addProperty("content", url);
//        vars1.add(infoperson0);
//        JsonObject infoperson1 = new JsonObject();
//        infoperson1.addProperty("name", "PERIOD1");
//        infoperson1.addProperty("content", dateStart);
//        vars1.add(infoperson1);
//        JsonObject infoperson2 = new JsonObject();
//        infoperson2.addProperty("name", "PERIOD2");
//        infoperson2.addProperty("content", dateEnd);
//        vars1.add(infoperson2);
//        JsonObject infoperson3 = new JsonObject();
//        infoperson3.addProperty("name", "NAMEREPORT");
//        infoperson3.addProperty("content", name);
//        vars1.add(infoperson3);
//        JsonObject var_mail1 = new JsonObject();
//        //var_mail.addProperty("rcpt", "ejimenez@moovin.me");
//        var_mail1.addProperty("rcpt", "Katty.pastor@moovin.me");
//        // PrintLog.PrintProduction("vars " + vars1);
//        var_mail1.add("vars", vars1);
//        merge_vars.add(var_mail1);
//        return merge_vars;
//
//    }

}
