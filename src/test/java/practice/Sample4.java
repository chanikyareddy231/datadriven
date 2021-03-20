package practice;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Sample4 {

	public static void main(String[] args) throws Exception
	{
		//open an exsting excel file(.xlsx) in read mode
		File f=new File("target\\dummy4.xlsx");
		FileInputStream fi=new FileInputStream(f);
		Workbook wb=WorkbookFactory.create(fi);
		Sheet sh=wb.getSheet("Sheet1");
        int nour=sh.getPhysicalNumberOfRows();
        int nouc=sh.getRow(0).getLastCellNum();
        //row sum
        for(int i=0;i<nour;i++) //row wise
        { 
        	   int  rowsum=0;
        	   for(int j=0;j<nouc;j++) //column wise in every row
        	   {
        		   int x=(int) sh.getRow(i).getCell(j).getNumericCellValue();
        		   rowsum=rowsum+x;
        	   }
        	   sh.getRow(i).createCell(nouc).setCellValue(rowsum);
        }
        //column sum
        for(int i=0;i<nouc;i++) //column wise
        { 
        	   int  colsum=0;
        	   for(int j=0;j<nour;j++) //row wise in each column
        	   {
        		   int x=(int) sh.getRow(j).getCell(i).getNumericCellValue();
        		   colsum=colsum+x;
        	   }
        	   if(i==0)
        	   {
        		   sh.createRow(nour).createCell(i).setCellValue(colsum);
        	   }
        	   else
        	   {
        		   sh.getRow(nour).createCell(i).setCellValue(colsum);
        	   }
        	   sh.autoSizeColumn(i); //auto fit 
         }
        //Take write permission on that file to save in HDD
        FileOutputStream fo=new FileOutputStream(f);
		wb.write(fo);
		wb.close();
		fi.close();
		fo.close();
	}

}
