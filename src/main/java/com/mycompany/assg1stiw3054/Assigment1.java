/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.assg1stiw3054;

/**
 *
 * @author User
 */
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.DataFormatter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.Writer;
import java.util.Iterator;

public class Assigment1 {
    
     public static void main(String[] args)
    {
           Writer write = null;
          boolean LineOut = true;
             
           
        try {
            
            DataFormatter dataformat = new DataFormatter();

            FileInputStream excel = new FileInputStream("C:\\Users\\User\\Documents\\NetBeansProjects\\Assg1STIW3054\\Practicum-StudentSupervisorList.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(excel);
            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = sheet.iterator();
            
            File markdown = new File("C:\\Users\\User\\Documents\\NetBeansProjects\\Assg1STIW3054\\243055.md");
            write = new BufferedWriter(new FileWriter(markdown));
               
            while (iterator.hasNext()) {

                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();
                
                while (cellIterator.hasNext()) {

                    Cell cell = cellIterator.next();
                    String data = dataformat.formatCellValue(cell);
                   
                    System.out.print(data +"|");
                  
                    write.write(data +"|"); 

                }
                System.out.println();
                write.write("\n");
                if (LineOut==true){
                    write.write("---|---|---|---|\n");
                    LineOut=false;
                }
                

            }
        } 
        catch (FileNotFoundException e) {
        } 
        catch (IOException e) {
        }
        
        try {                 
            if (write != null) {                     
                write.close();                 
            }             
        } catch (IOException e) {                 
            e.printStackTrace();             
        } 
    }
}
