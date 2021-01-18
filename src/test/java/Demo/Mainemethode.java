package Demo;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Mainemethode {

	public static void main(String[] args) throws IOException
	{
      FileInputStream file = new FileInputStream("C:\\Users\\Lenovo\\Desktop\\testcase.xlsx"); //file path
      XSSFWorkbook workbook = new XSSFWorkbook (file);
      
      int sheets= workbook.getNumberOfSheets();
      
      for(int i=0; i<sheets; i++)
      {
    	if(workbook.getSheetName(i).equalsIgnoreCase("sarika")) // checking file name
    	{
    	 XSSFSheet sheet= workbook.getSheetAt(i); // access sheet
    	 
    	 //// identify testcase colum by scannig entire row
    	 
    	Iterator <Row> rows = sheet.iterator(); //move to each and every row
    	Row firstrow  =rows.next(); //access specific entire row
      Iterator<Cell> ce= firstrow.cellIterator(); //move to each and every cell
      int k=0;
     int coloumn = 0;
     int a;
       while(ce.hasNext()) //read each and ever cell value, hasnext is says next cell is present or not
       {
    	   Cell value =ce.next(); //storing value cheack next cell is present and moved
    	  if (value.getStringCellValue().equalsIgnoreCase("TestCases")) //compare stored value with given value is same
    	  {
    		  coloumn =k;
    	  }
    	   k++;
       }
       
       System.out.println(coloumn);
       
       while(rows.hasNext())
       {
    	   Row r=rows.next();
    	   if(r.getCell(coloumn).getStringCellValue().equalsIgnoreCase("Purchase"))
    	   {
    		 Iterator<Cell> cv= r.cellIterator();
    		 while(cv.hasNext())
    		  {
    			  System.out.println(cv.next().getStringCellValue());
    		  }
    	   }
       }
    	 
    	 
    	}
      }

	}

}
