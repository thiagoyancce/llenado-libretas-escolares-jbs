/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package recursos;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javafx.scene.control.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Admin
 */
class RegistroCompetencias {
    private String[][]notas;
    private String []datos;
    private int bimestre;
    final static  Coordenada BIMESTRE=new Coordenada(4,22);// bimestre(5:6,w(23) 
final static  Coordenada GRADO=new Coordenada(5,2);//grado y seccion(6,c(3))
final static  Coordenada AREA=new Coordenada(4,12);//Curso(5,(m:13))

final static  Coordenada NOTA_INICIAL=new Coordenada(9,31); //Notas (10:44,ab(28):ae(31))
final static  Coordenada NOTA_FINAL=new Coordenada(43,32); 
final static  int NUMERO_ALUMNOS=35;
final static  int NUMERO_NOTAS=2;//COMPETENCIAS TRANSVERSALES

    RegistroCompetencias(int bimestre){
        notas=new String[NUMERO_ALUMNOS][NUMERO_NOTAS];
        this.bimestre=bimestre;
    }
    public boolean procesarDatos(String directorio) {
       
            try {
                leerExcel(directorio);
                return false;
            } catch (IOException ex) {
                Logger.getLogger(RegistroCompetencias.class.getName()).log(Level.SEVERE, null, ex);
            }
        
        return false;
    }

    private void leerExcel(String elemento) throws IOException {
        FileInputStream file;
        try {
            file = new FileInputStream(new File(elemento));
            XSSFWorkbook wb=new XSSFWorkbook(file);
            XSSFSheet sheet =wb.getSheetAt(bimestre); 
            String bimestre=sheet.getRow(BIMESTRE.getX()).getCell(BIMESTRE.getY()).getStringCellValue();
            String area=sheet.getRow(AREA.getX()).getCell(AREA.getY()).getStringCellValue();
            String grado=sheet.getRow(GRADO.getX()).getCell(GRADO.getY()).getStringCellValue();
            
          
            System.out.println("\n \n leyendo "+bimestre+" "+area+" "+grado);
            datos=new String []{String.valueOf(this.bimestre),area,grado};
            /*
            
            }
            
            */
            boolean esQuinto=grado.charAt(0)=='5';
            
            
                for (int i = NOTA_INICIAL.getX(); i <= NOTA_FINAL.getX(); i++) {
                   
                    for (int j = NOTA_INICIAL.getY(); j <= NOTA_FINAL.getY(); j++) {
                        
                        
                       switch(sheet.getRow(i).getCell(j).getCellTypeEnum().toString() ){
                       case "NUMERIC":
                         
                           notas[i-NOTA_INICIAL.getX()][j-NOTA_INICIAL.getY()]=String.valueOf((int)sheet.getRow(i).getCell(j).getNumericCellValue());
                           break;
                       case "STRING":
                        
                           notas[i-NOTA_INICIAL.getX()][j-NOTA_INICIAL.getY()]=sheet.getRow(i).getCell(j).getStringCellValue();
                           break; 
                        case "FORMULA":
                           
                            
                                notas[i-NOTA_INICIAL.getX()][j-NOTA_INICIAL.getY()]=sheet.getRow(i).getCell(j).getStringCellValue();
                            break;   
                        case "BLANK":
                           
                            notas[i-NOTA_INICIAL.getX()][j-NOTA_INICIAL.getY()]="";
                            break; 
                       }
                               
                    }   
                }
               /*
                    
                }
            */
           
        } catch (FileNotFoundException ex) {
            Logger.getLogger(RegistroCompetencias.class.getName()).log(Level.SEVERE, null, ex);
        }
     
     
    }
    
    String[][] recogerNotas(){
        return notas;
    }

    String[] recogerDatos() {
        return datos;
    }
}



class Coordenada{
   private  int x;
   private int y;
   public Coordenada(int x,int y){
       this.x=x;
       this.y=y;
   }
   public int  getX(){
       return x;
   }
   public int getY(){
       return y;
   }
}
