/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package recursos;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Admin
 */
class Consolidado {
   
    
   private String rutaSalida;
 final static Coordenada DPCC=new Coordenada(3,2);
 final static Coordenada CCSS=new Coordenada(3,4);
 final static Coordenada EDUCACIÓN_FISICA=new Coordenada(3,7);
 final static Coordenada ARTE_Y_CULTURA=new Coordenada(3,10);
 final static Coordenada COMUNICACIÓN=new Coordenada(3,12);
 final static Coordenada INGLÉS=new Coordenada(3,15);
 final static Coordenada MATEMÁTICA=new Coordenada(3,18);
 final static Coordenada CIENCIA_Y_TECNOLOGÍA=new Coordenada(3,22);
 final static Coordenada EDUCACIÓN_RELIGIOSA=new Coordenada(3,25);
  final static Coordenada EPT=new Coordenada(3,27);
   final static String NOMBRE_ARCHIVO="\\CONSOLIDADO-LIT-2022.xlsx";
    Consolidado(String rutaSalida){
       // this.plantilla=plantilla;
        this.rutaSalida=rutaSalida;
        }
    void procesarDatos(String[][] notas, String[] bag) throws IOException {
     
        abrirPlantilla( notas,bag);
    
    }

    private void abrirPlantilla(String[][] notas,String[] bag) throws IOException {
        
         Coordenada c=selectCoordenadaCurso(bag[1]);
         int numeroNotas=selectTamañoCurso(bag[1]);
        FileInputStream file;
        
            System.out.println(" escribiendo archivo literal");
           // System.out.println(rutaSalida+"\\Consolidado"+bag[2]+".xlsx" );
           
            
       
        if(new File(rutaSalida+"\\Consolidado"+bag[2]+".xlsx").isFile()){
                //System.out.println(" existe "+bag[2]);
            file=new FileInputStream(new File(rutaSalida+"\\Consolidado"+bag[2]+".xlsx"));
        }
        else
            file=new FileInputStream(new File(rutaSalida+NOMBRE_ARCHIVO));
        
        try {
            
            XSSFWorkbook wb=new XSSFWorkbook(file);
            XSSFSheet sheet =wb.getSheetAt(Integer.parseInt(bag[0]));
            System.out.print("escribiendo...");
            sheet.getRow(0).getCell(2).setCellValue(bag[2]); 
          
            for (int i = c.getX(); i < c.getX()+35; i++) {
                    for (int j = c.getY(); j < c.getY()+numeroNotas; j++) {
                        //System.out.println("esto...."+notas[i-c.getX()][j-c.getY()]);
                        if(!notas[i-c.getX()][j-c.getY()].equals(""))
                           sheet.getRow(i).getCell(j).setCellValue(notas[i-c.getX()][j-c.getY()]);                     
                        
                        }   
                }
            
            file.close();
            FileOutputStream output= new FileOutputStream(rutaSalida+"\\Consolidado"+bag[2]+".xlsx");
            wb.write(output);
            output.close();
            System.out.print("...Fin");
               
        } catch (FileNotFoundException ex) {
            Logger.getLogger(Consolidado.class.getName()).log(Level.SEVERE, null, ex);
        }
     
     
    
    }

    private Coordenada selectCoordenadaCurso(String curso) {
 
        switch(curso){
            case "DPCC":return DPCC;
            case "CIENCIAS SOCIALES":return CCSS;  
            case "EDUCACIÓN FÍSICA":return EDUCACIÓN_FISICA;
            case "ARTE Y CULTURA":return ARTE_Y_CULTURA;
            case "COMUNICACIÓN":return COMUNICACIÓN;
            case "INGLÉS":return INGLÉS;
            case "MATEMÁTICA":return MATEMÁTICA;
            case "CIENCIA Y TECNOLOGÍA":return CIENCIA_Y_TECNOLOGÍA;
            case "EDUCACIÓN RELIGIOSA": return  EDUCACIÓN_RELIGIOSA;
            case "EPT":return EPT;
            default :System.out.println(curso+" no es un curso");
                  return new Coordenada(0,0);
        }
        
    }

    private int seleccionarBimestre(String bimestre) {
        switch(bimestre){
            case "0": return 0;
            case "1": return 1;
            case "2": return 2;
            case "3": return 3;
            default:return -1;
            
        
        }// generated methods, choose Tools | Templates.
    }

    private int selectTamañoCurso(String curso) {
         switch(curso){
            case "DPCC":return 2;
            case "CIENCIAS SOCIALES":return 3;          
            case "EDUCACIÓN FÍSICA":return 3;
            case "ARTE Y CULTURA":return 2;
            case "COMUNICACIÓN":return 3;
            case "INGLÉS":return 3;
            case "MATEMÁTICA":return 4;
            case "CIENCIA Y TECNOLOGÍA":return 3;
            case "EDUCACIÓN RELIGIOSA":return 2;
            case "EPT":return 1;
         
            default :System.out.println(curso+" N es un curso");
                  return 0;
        }
    
    }
}