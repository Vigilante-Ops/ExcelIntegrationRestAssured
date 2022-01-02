import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class dataDriven {

    public static ArrayList<Object> GetData(String testCaseName) throws IOException {
        ArrayList<Object> a= new ArrayList<Object>();
        FileInputStream fis=new FileInputStream("C:\\Users\\Aman\\IdeaProjects\\RestAssuredExcel\\src\\utils\\Book2.xlsx");
        XSSFWorkbook workbook=new XSSFWorkbook(fis);
        int sheets=workbook.getNumberOfSheets();
        int k=0;
        int column=0;
        for (int i=0;i<sheets;i++)
        {
            if(workbook.getSheetName(i).equalsIgnoreCase("testData"))
            {
                //get sheet
                XSSFSheet sheet= workbook.getSheetAt(i);
                Iterator<Row> rows= sheet.iterator();//sheet is collection of row
                Row firstRow=  rows.next();
                Iterator<Cell> ce= firstRow.cellIterator();//row is collection of cells
                while (ce.hasNext())
                {
                    Cell value= ce.next();
                    if(value.getStringCellValue().equalsIgnoreCase("Testcases"))
                    {
                        column=k;
                    }
                    k++;
                }

                while (rows.hasNext())
                {
                    Row r= rows.next();
                    if(r.getCell(column).getStringCellValue().equalsIgnoreCase(testCaseName))
                    {
                        Iterator<Cell> cv=r.cellIterator();
                        while (cv.hasNext())
                        {
                            Cell c=cv.next();
                            if(c.getCellType()== CellType.STRING) {
                                a.add(c.getStringCellValue());
                            }
                            else
                            {
                                a.add(c.getNumericCellValue());

                            }
                        }



                    }


                }

            }
        }
        return a;


    }


    public static void main(String[] args) throws IOException {

        System.out.println(GetData("Purchase"));


    }
}
