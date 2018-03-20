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
    
     public static void main(String[] args) {

        Writer writer = null;
        boolean LineOut = true;

        try {

            DataFormatter dataformat = new DataFormatter();
            FileInputStream excel = new FileInputStream(new File("C:\\Users\\User\\Documents\\NetBeansProjects\\Assg1STIW3054\\Practicum-StudentSupervisorList.xlsx"));
            Workbook workbook = new XSSFWorkbook(excel);
            Sheet data = workbook.getSheetAt(0);
            Iterator<Row> iterator = data.iterator();
            
            File markdown = new File("C:\\Users\\User\\Documents\\NetBeansProjects\\Assg1STIW3054\\243055.md");
            writer = new BufferedWriter(new FileWriter(markdown));

            while (iterator.hasNext()) {

                Row row = iterator.next();
                Iterator<Cell> cellIterator = row.iterator();

                while (cellIterator.hasNext()) {

                    Cell cell = cellIterator.next();
                    String value = dataformat.formatCellValue(cell);

                    System.out.print(value + "|");

                    writer.write(value + "|");

                }
                System.out.println();
                writer.write("\n");
                if (LineOut == true) {
                    writer.write("---|---|---|---|\n");
                    LineOut = false;
                }

            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        try {
            if (writer != null) {
                writer.close();
            }
        } catch (IOException e) {
            e.printStackTrace();
        } 
    }
}