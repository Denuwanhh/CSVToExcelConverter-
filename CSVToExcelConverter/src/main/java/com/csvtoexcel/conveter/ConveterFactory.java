/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package com.csvtoexcel.conveter;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import org.apache.log4j.Logger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Denuwan
 */
public class ConveterFactory {
    
    private static Logger logger = Logger.getLogger(ConveterFactory.class);

    public static void convertCsvToExcel(File inputFile, File outputFile) {

        StringBuilder stringBuilderValue;
        FileOutputStream fileOutputStream = null; 
        XSSFWorkbook xssfWorkbook = null;
        try {
            stringBuilderValue = new StringBuilder();
            fileOutputStream = new FileOutputStream(outputFile);
            xssfWorkbook = new XSSFWorkbook(new FileInputStream(inputFile));
            Iterator<Sheet> sheets = xssfWorkbook.iterator();

            while (sheets.hasNext()) {
                Cell cell;
                Row row;
                Sheet sheet = sheets.next();
                Iterator<Row> rowIterator = sheet.iterator();
                
                while (rowIterator.hasNext()) {                    
                    row = rowIterator.next();
                    Iterator<Cell> cellIterator = row.cellIterator();
                    
                    while (cellIterator.hasNext()) {
                        cell = cellIterator.next();

                        switch (cell.getCellType()) {
                            case Cell.CELL_TYPE_BOOLEAN:
                                stringBuilderValue.append(cell.getBooleanCellValue()).append(",");
                                break;

                            case Cell.CELL_TYPE_NUMERIC:
                                stringBuilderValue.append(cell.getNumericCellValue()).append(",");
                                break;

                            case Cell.CELL_TYPE_STRING:
                                stringBuilderValue.append(cell.getStringCellValue()).append(",");
                                break;

                            case Cell.CELL_TYPE_BLANK:
                                stringBuilderValue.append("" + ",");
                                break;

                            default:
                                stringBuilderValue.append(cell).append(",");
                        }
                    }
                    stringBuilderValue.append('\n');
                }
                fileOutputStream.write(stringBuilderValue.toString().getBytes());
            }
           
            logger.info("File Convert Successfully !");
            
        } catch (FileNotFoundException ex) {
            logger.error("Exception In convertCsvToExcel() Method?=  " + ex);
        } catch (IOException ex) {
           logger.error("Exception In convertCsvToExcel() Method?=  " + ex);
        }finally{
            try{
                fileOutputStream.close();
                xssfWorkbook.close();
            }catch(IOException ex){
                logger.error("Exception In convertCsvToExcel() Method?=  " + ex);
            }
        }
    }
}
