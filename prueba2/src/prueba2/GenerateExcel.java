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
        String[] headerBD = headers;
        
        String[] headerSIEBEL = new String[]{ "FechaSiebel", "Cedula", "Nombre",
                                            "Telefono", "# PEDIDO SIEBEL", "TIPO DE PLAN",
                                            "Validacion DE PLAN", "Canal", "Realizado POR",
                                            "Confirmacion", "META DE MES", "REBAJA ICE",
                                            "MONTO DE PLAN", "% COMISION", "MONTO COMISION"
                                            };
        
        try {
            // --------------------- instancia de excel --------------------- 
            HSSFWorkbook workbook = new HSSFWorkbook();         // Libros 
            HSSFSheet sheet, sheet2, sheet3 ;  
           
            // ------------------- estilo del encabezado -------------------
            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.SEA_GREEN.getIndex());  // agregarle un color al fondo de las celdas.
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setAlignment(HorizontalAlignment.CENTER);
            
            sheet =  createSheetSiebel( workbook, headerStyle, headerSIEBEL,  dataInfo, "Siebel");
            sheet2 = createSheetBD(     workbook, headerStyle, headerBD,      dataInfo, "BD"); 
            sheet3 = createSheetPrueba( workbook, headerStyle, headers,       dataInfo, "Prueba);

            // ------------------- se almacena archivo excel -------------------
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
      
    public HSSFSheet createSheetBD(HSSFWorkbook pWorkbook, CellStyle pHeaderStyle, String[] headers, ArrayList<ArrayList> dataInfo, String pSheetname){
        HSSFSheet sheet = pWorkbook.createSheet(); 
        pWorkbook.setSheetName(1, pSheetname);                 // nombres de hojas
        int positionHeader = 0 ;
        int initialPosititionData = positionHeader + 1;
         
        // Estilos 
        Font font = pWorkbook.createFont();      
        font.setBold(true);
        pHeaderStyle.setFont(font);
        
        // ----------------- ANADIMOS DATOS AL HEADER ---------------------- 
        HSSFRow headerRow = sheet.createRow(positionHeader);  // se crea una fila en la hoja en la posicion 0.
        // creamos el encabezado
        for (int i = 0; i < headers.length; ++i) {
            String header = headers[i];
            // creamos celda en la fila creada anterior en la posicion del for.
            HSSFCell cell = headerRow.createCell(i);
            // le aplicamos el estilo
            cell.setCellStyle(pHeaderStyle);
            cell.setCellValue(header);
        } 
        
        for (int i = 0; i < dataInfo.size(); ++i) { 
                ArrayList<String> row = dataInfo.get(i); 
                // creamos una fila nueva para los datos anteriores.
                HSSFRow dataRow = sheet.createRow(i + initialPosititionData );
                String formula ;
                System.out.print("tamano fila: "+ row.size());
                for (int j = 0; j < row.size(); j++) { 
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
                    sheet.autoSizeColumn(j);        // auto ajustador de columnas
                }   
            } 
        return sheet; 
    } 
    
    public HSSFSheet createSheetSiebel(HSSFWorkbook pWorkbook, CellStyle pHeaderStyle, String[] headers, ArrayList<ArrayList> dataInfo, String pSheetname){
        HSSFSheet sheet = pWorkbook.createSheet(); 
        pWorkbook.setSheetName(0, pSheetname);                 // nombres de hojas
        int positionHeader = 0 ;
        int initialPosititionData = positionHeader + 1; 
        Font font = pWorkbook.createFont();                     // Estilos  
        font.setBold(true);
        pHeaderStyle.setFont(font);
        
        // ----------------- ANADIMOS DATOS AL HEADER ---------------------- 
        HSSFRow headerRow = sheet.createRow(positionHeader);    // se crea una fila en la hoja en la posicion 0.
        // creamos el encabezado
        for (int i = 0; i < headers.length; ++i) {
            String header = headers[i];
            // creamos celda en la fila creada anterior en la posicion del for.
            HSSFCell cell = headerRow.createCell(i); 
            cell.setCellStyle(pHeaderStyle);                    // le aplicamos el estilo
            cell.setCellValue(header);
        } 
        
        for (int i = 0; i < dataInfo.size(); ++i) { 
                ArrayList<String> row = dataInfo.get(i); 
                // creamos una fila nueva para los datos anteriores.
                HSSFRow dataRow = sheet.createRow(i + initialPosititionData );
                String formula ;
                System.out.print("tamano fila: "+ row.size());
                for (int j = 0; j < row.size(); j++) { 
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
                            formula = "(M"+col+ "-(M" +col+ "* L"+col+"/100 ) )* N"+col+ "/100"; //  =(H6-(H6*11,95%))*10% 
                            System.out.println(formula);
                            dataRow.createCell(j).setCellFormula(formula );  
                            break; 
                         default: 
                            dataRow.createCell(j).setCellValue(row.get(j));  
                            break; 
                    }  
                    sheet.autoSizeColumn(j);                    // auto ajustador de columnas
                }   
            } 
        return sheet; 
    }
    
    // Se comprueba que se puede realizar calculos con hojas vecinas del mismo libro
    public HSSFSheet createSheetPrueba(HSSFWorkbook pWorkbook, CellStyle pHeaderStyle, String[] headers, ArrayList<ArrayList> dataInfo, String pSheetname){
        HSSFSheet sheet = pWorkbook.createSheet(); 
        pWorkbook.setSheetName(0, pSheetname);                 // nombres de hojas
        int positionHeader = 0 ;
        int initialPosititionData = positionHeader + 1; 
        Font font = pWorkbook.createFont();                     // Estilos  
        font.setBold(true);
        pHeaderStyle.setFont(font);
        
        // ----------------- ANADIMOS DATOS AL HEADER ---------------------- 
        HSSFRow headerRow = sheet.createRow(positionHeader);    // se crea una fila en la hoja en la posicion 0.
        // creamos el encabezado
        for (int i = 0; i < headers.length; ++i) {
            String header = headers[i];
            // creamos celda en la fila creada anterior en la posicion del for.
            HSSFCell cell = headerRow.createCell(i); 
            cell.setCellStyle(pHeaderStyle);                    // le aplicamos el estilo
            cell.setCellValue(header);
        } 
        
        for (int i = 0; i < dataInfo.size(); ++i) { 
                ArrayList<String> row = dataInfo.get(i); 
                // creamos una fila nueva para los datos anteriores.
                HSSFRow dataRow = sheet.createRow(i + initialPosititionData );
                String formula ;
                System.out.print("tamano fila: "+ row.size());
                for (int j = 0; j < row.size(); j++) { 
                    int col = i+2 ; // donde inicia los datos a escribirse
                    switch(j){ 
                        case 10:         // columna O: MONTO COMISION
                            System.out.print("tamano comision: "+ row.get(j)); 
                            formula = "(M"+col+ "-(M" +col+ "* L"+col+"/100 ) )* N"+col+ "/100"; //  =(H6-(H6*11,95%))*10% 
                            System.out.println(formula);
                            dataRow.createCell(j).setCellFormula(formula );  
                            break; 
                         default: 
                            dataRow.createCell(j).setCellValue(row.get(j));  
                            break; 
                    }  
                    sheet.autoSizeColumn(j);                    // auto ajustador de columnas
                }   
            } 
        return sheet; 
    }
    
    
    
}
