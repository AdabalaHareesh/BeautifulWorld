import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class Writesheet {

	@SuppressWarnings("resource")
	public static void main(String[] args) throws Exception 
   {
      //Create blank workbook
      XSSFWorkbook workbook = new XSSFWorkbook(); 
      //Create a blank sheet
      XSSFSheet spreadsheet = workbook.createSheet( 
      " Students No ");
      //Create row object
      XSSFRow row;
      //This data needs to be written (Object[])
      Map < String, Object[] > stdntno = 
      new TreeMap < String, Object[] >();
      stdntno.put( "1", new Object[] { 
      "NAME", "Phone No" });
      stdntno.put( "2", new Object[] { "Gopal","9505068767" });
      stdntno.put( "3", new Object[] { "Manis","9985999966" });
      stdntno.put( "4", new Object[] { "Masth","9494022013" });
      stdntno.put( "5", new Object[] { "Satis","9490745168" });
      stdntno.put( "6", new Object[] { "Krish","9666776326" });
      //Iterate over data and write to sheet
      Set < String > keyid = stdntno.keySet();
      int rowid = 0;
      for (String key : keyid)
      {
         row = spreadsheet.createRow(rowid++);
         Object [] objectArr = stdntno.get(key);
         int cellid = 0;
         for (Object obj : objectArr)
         {
            Cell cell = row.createCell(cellid++);
            cell.setCellValue((String)obj);
         }
      }
      //Write the workbook in file system
      FileOutputStream out = new FileOutputStream( 
      new File("Writesheet.xlsx"));
      workbook.write(out);
      out.close();
      System.out.println( 
      "Writesheet.xlsx written successfully" );
   }
}