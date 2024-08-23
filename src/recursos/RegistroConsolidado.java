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
class RegistroConsolidado {
    private String[][]notas;
    private String []datos;
    private int bimestre;
    final static  Coordenada BIMESTRE=new Coordenada(4,22);// bimestre(5:6,w(23) 
final static  Coordenada GRADO=new Coordenada(5,2);//grado y seccion(6,c(3))
final static  Coordenada AREA=new Coordenada(4,12);//Curso(5,(m:13))

final static  Coordenada NOTA_INICIAL=new Coordenada(9,27); //Notas (10:44,ab(28):ae(31))
final static  Coordenada NOTA_FINAL=new Coordenada(43,30); //Notas (10:44,ab(28):ae(31))
    
 
final static  int NUMERO_ALUMNOS=35;
final static  int NUMERO_NOTAS=4;//COMPETENCIAS TRANSVERSALES

    RegistroConsolidado(int bimestre){
        notas=new String[NUMERO_ALUMNOS][NUMERO_NOTAS];
        this.bimestre=bimestre;
    }
    public boolean procesarDatos(String directorio) {
       
            try {
                leerExcel(directorio);
                return false;
            } catch (IOException ex) {
                Logger.getLogger(RegistroConsolidado.class.getName()).log(Level.SEVERE, null, ex);
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
            int totalAlumnos=0;
            for (int i = 9; i < 40; i++) {
                if(!sheet.getRow(i).getCell(1).getStringCellValue().equals("0"));
                    totalAlumnos++;
            }
            
            */
            boolean esQuinto=grado.charAt(0)=='5';
            
            
                for (int i = NOTA_INICIAL.getX(); i <= NOTA_FINAL.getX(); i++) {
                    for (int j = NOTA_INICIAL.getY(); j <= NOTA_FINAL.getY(); j++) {
                        
                        
                       switch(sheet.getRow(i).getCell(j).getCellTypeEnum().toString() ){
                       case "NUMERIC":
                         
                           notas[i-NOTA_INICIAL.getX()][j-NOTA_INICIAL.getY()]=String.valueOf(sheet.getRow(i).getCell(j).getNumericCellValue());
                           break;
                       case "STRING":
                        
                           notas[i-NOTA_INICIAL.getX()][j-NOTA_INICIAL.getY()]=sheet.getRow(i).getCell(j).getStringCellValue();
                           break; 
                        case "FORMULA":
                            notas[i-NOTA_INICIAL.getX()][j-NOTA_INICIAL.getY()]=sheet.getRow(i).getCell(j).getStringCellValue();
                            break;   
                        case "BLANK":
                            notas[i-NOTA_INICIAL.getX()][j-NOTA_INICIAL.getY()]="";
                           
                       }
                               
                    }   
                }
               /* for (int i = 0; i < notas.length; i++) {
                    for (int j = 0; j < notas[1].length; j++) {
                        if(notas[i][j]==null)
                            notas[i][j]="0";
                    }   
                    
                }
            */
           
        } catch (FileNotFoundException ex) {
            Logger.getLogger(RegistroConsolidado.class.getName()).log(Level.SEVERE, null, ex);
        }
     
     
    }
    
    String[][] recogerNotas(){
        return notas;
    }

    String[] recogerDatos() {
        return datos;
    }
}



