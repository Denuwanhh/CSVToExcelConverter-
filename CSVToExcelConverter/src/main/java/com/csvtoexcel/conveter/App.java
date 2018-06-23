package com.csvtoexcel.conveter;

import java.io.File;

/**
 * Execution class
 * @author Denuwan
 */
public class App 
{
    public static void main( String[] args ){
        File inputFile = new File("C:\\Users\\Denuwan\\Desktop\\test\\test.csv");
        File outputFile = new File("C:\\Users\\Denuwan\\Desktop\\test\\test.xls");
        ConveterFactory.convertCsvToExcel(inputFile, outputFile);
    }
}
