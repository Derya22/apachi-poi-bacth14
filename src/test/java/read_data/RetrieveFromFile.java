package read_data;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

 // we get data from very first row and cell
public class RetrieveFromFile {

   @Test
   public void readFileTest() throws IOException {

     File excelFile = new File("src/test/resources/TestSetup.xlsx");

     FileInputStream fileInputStream = new FileInputStream(excelFile);
     //we will pass this stream to apachi-poi library that library will be sorting out the data
     // for us and using those data and retrieve

     XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);// workbook means whole Excel file.
     XSSFSheet page1 = workbook.getSheet("Sheet1"); // we are getting specific one sheet from the Excel document we were working on
     XSSFRow row1 = page1.getRow(0); // we use indexing it starts from 0,it will give first row
     XSSFCell cell1 = row1.getCell(0);

     System.out.println(cell1);

     XSSFRow row2 = page1.getRow(1);
     XSSFCell cell2 = row2.getCell(0);
     System.out.println(cell2);


   }

   @Test

   public void getRowValuesTest() throws IOException {

     File file = new File("src/test/resources/TestSetup.xlsx");
     FileInputStream fileInputStream = new FileInputStream(file);

     XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
     XSSFSheet sheet1 = workbook.getSheetAt(0);
     XSSFRow row1 = sheet1.getRow(0);

     for (int i = row1.getFirstCellNum(); i < row1.getLastCellNum(); i++) { //getfirstCellNum is very first index.
       //if there is no value it won't take it.
       XSSFCell cell = row1.getCell(i);
       System.out.print(cell + " | ");

     }

     }


   @Test
   public void getAllDataTest() throws IOException {

     //GET all data from excel document
     File file = new File("src/test/resources/TestSetup.xlsx");
     FileInputStream fileInputStream = new FileInputStream(file);
     XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
     XSSFSheet sheet1 = workbook.getSheetAt(0);


     for (int i = sheet1.getFirstRowNum(); i < sheet1.getLastRowNum(); i++) {
       XSSFRow tempRow = sheet1.getRow(i);
       System.out.print("| "); // adds pipe to beginning
       for (int j = tempRow.getFirstCellNum(); j < tempRow.getLastCellNum(); j++) {
         System.out.print(tempRow.getCell(j) + " | ");//it prints everything same line
       }
       System.out.println();//new line print out.
     }


   }

   }



