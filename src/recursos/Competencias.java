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
class Competencias {
    
    
    String rutaSalida;
 final static Coordenada DPCC=new Coordenada(3,2);
 final static Coordenada CCSS=new Coordenada(3,4);
 final static Coordenada EDUCACIÓN_FISICA=new Coordenada(3,6);
 final static Coordenada ARTE_Y_CULTURA=new Coordenada(3,8);
 final static Coordenada COMUNICACIÓN=new Coordenada(3,10);
 final static Coordenada INGLÉS=new Coordenada(3,12);
 final static Coordenada MATEMÁTICA=new Coordenada(3,14);
 final static Coordenada CIENCIA_Y_TECNOLOGÍA=new Coordenada(3,16);
 final static Coordenada EDUCACIÓN_RELIGIOSA=new Coordenada(3,18);
  final static Coordenada EPT=new Coordenada(3,20);
   final static Coordenada COMPETENCIA_TRANSVERSALES=new Coordenada(3,22);
   static String[] RUTA_PLANTILLA_COMP={"\\COMP-TRANS-1RO.xlsx","\\COMP-TRANS-2DO.xlsx","\\COMP-TRANS-3RO.xlsx","\\COMP-TRANS-4TO.xlsx","\\COMP-TRANS-5TO.xlsx"};
  
    Competencias(String rutaSalida){
       // this.plantilla=plantilla;
        this.rutaSalida=rutaSalida;
       
     }
    void procesarDatos(String[][] notas, String[] bag) throws IOException {
     
        abrirPlantilla( notas,bag);
    
    }

    private void abrirPlantilla(String[][] notas,String[] bag) throws IOException {
        int hoja=seleccionarHoja(bag[2].charAt(1));
         Coordenada c=selectCoordenadaCurso(bag[1]);
        // int tamaño=selectTamañoCurso(bag[1]);
         int tamaño=RegistroCompetencias.NUMERO_NOTAS;
        FileInputStream file;
        String rutaFinal="";
        System.out.println( "grado "+bag[2].charAt(0)+bag[2].charAt(1));
        switch (bag[2].charAt(0)) {
            case '1':
                rutaFinal=rutaSalida+RUTA_PLANTILLA_COMP[0];
                break;
            case '2':
                rutaFinal=rutaSalida+RUTA_PLANTILLA_COMP[1];
                break;
            case '3':
                rutaFinal=rutaSalida+RUTA_PLANTILLA_COMP[2];
                break;
            case '4':
                rutaFinal=rutaSalida+RUTA_PLANTILLA_COMP[3];
                break;
            case '5':
                rutaFinal=rutaSalida+RUTA_PLANTILLA_COMP[4];
                break;
            default:
                System.out.println("error leer grado");
                break;
        }
       if(new File(rutaFinal).isFile()){
                System.out.println(" seccion "+bag[2]);
                file=new FileInputStream(new File(rutaFinal));
            
                try {
            
                    XSSFWorkbook wb=new XSSFWorkbook(file);
                    XSSFSheet sheet =wb.getSheetAt(hoja);
                    System.out.print("escribiendo...");
                    for (int i = c.getX(); i < c.getX()+35; i++) {
                        for (int j = c.getY(); j < c.getY()+tamaño; j++) {
                        //System.out.println("esto...."+notas[i-c.getX()][j-c.getY()]);
                            if(!notas[i-c.getX()][j-c.getY()].equals("")){                     
                                sheet.getRow(i).getCell(j).setCellValue(notas[i-c.getX()][j-c.getY()]);                     
                            }
                        }   
                    }
            
                    file.close();
                    FileOutputStream output= new FileOutputStream(rutaFinal);
                    wb.write(output);
                    output.close();
                    
               
        } catch (FileNotFoundException ex) {
            Logger.getLogger(Competencias.class.getName()).log(Level.SEVERE, null, ex);
        }
       
       
       
       }
       else{
           System.out.println("error archivos no existe");
           
        
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

    private int seleccionarHoja(char seccion) {
        switch(seccion){
            case 'A': return 0;
            case 'B': return 1;
            case 'C': return 2;
            case 'D': return 3;
            case 'E': return 4;
            case 'F': return 5;
            case 'G': return 6;
            case 'H': return 7;
            case 'I': return 8;
            case 'J': return 9;
            case 'K': return 10;
            default  : return  -1;
            
        
        }// generated methods, choose Tools | Templates.
    }

    
    
    
}