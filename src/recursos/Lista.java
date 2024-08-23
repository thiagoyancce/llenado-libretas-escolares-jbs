/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package recursos;

import java.util.ArrayList;

/**
 *
 * @author Admin
 */
class Lista {
    private ArrayList <String>directorios=new ArrayList();
    Lista(){
       
    }
    void agregarDirectorios(String ruta,String[] listaDirectorios){
        for (int i = 0; i < listaDirectorios.length; i++) {
            if(listaDirectorios[i].contains(".xlsx")&&!listaDirectorios[i].contains("~$"))
            directorios.add(ruta+"\\"+listaDirectorios[i]);
            
        }
        
    }
    public void vaciarDirectorios(){
        for (int i = 0; i < directorios.size(); i++) {
            eliminarDirectorio(i);
            
        }
        
    }
    public ArrayList darLista(){
        return directorios;
    }
    void mostrarLista(){
      
            System.out.println(directorios);
        
    }
    String  darElemento(int index){
        return directorios.get(index).toString();
    }
    void eliminarDirectorio(int posicion){    
    
            directorios.remove(posicion);
       
       
    }
    void agregarDirectorio(String directorio){
        directorios.add(directorio);          
    }
    int darTamaño(){
    return directorios.size();
    }
    
}
