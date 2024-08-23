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
class Prueba {
   
    private String []resultados;
    private int bimestre;
    private String leyenda;
    final static  Coordenada BIMESTRE=new Coordenada(4,22);// bimestre(5:6,w(23) 
final static  Coordenada GRADO=new Coordenada(5,2);//grado y seccion(6,c(3))
final static  Coordenada AREA=new Coordenada(4,12);//Curso(5,(m:13))

final static  Coordenada NOTA_INICIAL=new Coordenada(9,31); //Notas (10:44,ab(28):ae(31))
final static  Coordenada NOTA_FINAL=new Coordenada(43,32); 


    Prueba(int bimestre){
  
        this.bimestre=bimestre;
    }
   

    private String leerExcel(String elemento) throws IOException {
        FileInputStream file;
        try {
            file = new FileInputStream(new File(elemento));
            XSSFWorkbook wb=new XSSFWorkbook(file);
            XSSFSheet sheet =wb.getSheetAt(bimestre); 
            
            String area=sheet.getRow(AREA.getX()).getCell(AREA.getY()).getStringCellValue();
            String grado=sheet.getRow(GRADO.getX()).getCell(GRADO.getY()).getStringCellValue();
            leyenda=sheet.getRow(45).getCell(1).getStringCellValue();
            wb.close();
            file.close();
           
           return (bimestre+1)+" Bimestre "+area+" "+grado+" "+alumnosCompletos()+"\n ";
           
        } catch (FileNotFoundException ex) {
           
            Logger.getLogger(RegistroCompetencias.class.getName()).log(Level.SEVERE, null, ex);
        return "error ";
        }
     
     
    }
    
    String[]resultado(){
        return resultados;
    }
    String alumnosCompletos(){
    if(leyenda.equals("LEYENDA:"))
        return "completo";
    else
        return "incompleto";
    }
    

  

    void revisar(Lista listaDirectory) {
        resultados=new String[listaDirectory.darTamaño()];
        try {
            for (int i = 0; i < listaDirectory.darTamaño(); i++) {
                resultados[i]="archivo "+(i+1)+") "+leerExcel(listaDirectory.darElemento(i));
                
            }
               
                
            } catch (IOException ex) {
                Logger.getLogger(Prueba.class.getName()).log(Level.SEVERE, null, ex);
            }
        
    }
}




