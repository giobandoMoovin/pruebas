/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package prueba2;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;

/**
 *
 * @author gil
 */
public class Prueba2 {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        // TODO code application logic here 
        crearExcelBasico("2018-01-01","2019-01-01"); 
    }
    
    public static void crearExcelBasico(String dateStart, String dateEnd){
        String fileURL = "";
//        Connection connection = this.connector.getConnection();
        String[] headers = new String[]{
            "FECHA",                "CEDULA",           "NOMBRE",
            "TELEFONO",             "# PEDIDO SIEBEL",  "TIPO DE PLAN", 
            "VALIDACION DE PLAN",   "CANAL",            "REALIZADO POR", 
            "CONFIRMACION",         "META DE MES",      "REBAJA ICE",       
            "MONTO DE PLAN",        "% COMISION",       "MONTO COMISION"
        };
        
        ArrayList<ArrayList> dataInfo = new ArrayList<ArrayList>();
        
        try {
//            PrintLog.Print(query);
//            Statement stmt = connection.createStatement();
//            ResultSet result = stmt.executeQuery(query);
//            while (result.next()) { 
            for(int i =0; i< 10; i++){
                
                // TENER CUIDADO PORQUE SI SON DECIMAS DEBE SER CON COMAS.
                ArrayList<String> row = new ArrayList<>();
                row.add("2/ener/2019");         //1 FECHA
                row.add("30432032"+i);          //2 CEDULA
                row.add("fulanit de jesus");    //3 NOMBRE
                row.add("888888888");           //4 TELEFONO
                row.add(i + "-42421253");       //5 # PEDIDO
                row.add("Plan 4g K3 24 M");     //6 TIPO DE PLAN
                row.add("Nuevo Plan 4g K3 24 M");//7 VALIDACION DE PLAN
                row.add("tlmk");                //8 CANAL
                row.add("Leo Umana");           //9 REALIZADO POR
                row.add(i+i+"00");              //10 CONFIRMACION
                row.add("138");                 //11 META DEL MES
                row.add("11,95");               //12 % ICE
                row.add("8000");                //13 MONTO PLAN
                row.add(""+(Integer.parseInt("10")+i));                  //14 % COMISION
                row.add("0000");                //15 MONTO COMISION 
                
                dataInfo.add(row);
            } 
//                row.add(result.getString("attentionChannel"));
//                row.add(result.getString("campaingName"));
//                row.add(result.getString("clientIdentification"));
//                row.add(result.getString("client"));
//                row.add(result.getString("executive"));
//                row.add(result.getString("orderNumber"));
//                row.add(result.getString("saleDate"));
//                row.add(result.getString("saleType"));
//                row.add(result.getString("nonSaleCause"));
//                row.add(result.getString("phone"));
//                dataInfo.add(row); 
//            }
//            result.close();
//            stmt.close(); 

            GenerateExcel generate = new GenerateExcel();
            fileURL = generate.saveExcel(headers, dataInfo, "Total de Planes", dateStart, dateEnd);
        }
        catch(Exception e){
            System.out.print("Error: "+ e);
            
        }
//        catch (SQLException ex) {
//            Logger.getLogger(Sale.class.getName()).log(Level.SEVERE, null, ex);
//        }
//        finally {
//            if (connection != null) {
//                try {
//                    connection.close();
//                } catch (SQLException ex) {
//                    Logger.getLogger(Sale.class.getName()).log(Level.SEVERE, null, ex);
//                }
//            }
//            connection = null;
//        }  
    } 
    
}
