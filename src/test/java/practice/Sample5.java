package practice;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Sample5 {

	public static void main(String[] args) throws Exception
	{
		//connect an existing excel file(.xlsx) in read mode
		File f=new File("target\\dummy5.xlsx");
		FileInputStream fi=new FileInputStream(f);
		Workbook wb=WorkbookFactory.create(fi);
		Sheet sh=wb.getSheet("Sheet1");
        int nour=sh.getPhysicalNumberOfRows();
        int nouc=sh.getRow(0).getLastCellNum();
        //row max
        for(int i=0;i<nour;i++)
        {
        	int rowmax=(int) sh.getRow(i).getCell(0).getNumericCellValue();
        	for(int j=1;j<nouc;j++) //Column wise in every row
        	{
				int x=(int) sh.getRow(i).getCell(j).getNumericCellValue();
        	    if(rowmax<x)
        	    {
        	    	rowmax=x;
        	    }
        	}
        	sh.getRow(i).createCell(nouc).setCellValue(rowmax);
        	}
        //colmax
        for(int i=0;i<nouc;i++) //colwise
        {
        	int colmax=(int) sh.getRow(i).getCell(0).getNumericCellValue();
        	for(int j=1;j<nour;j++) //Row wise in every row
        	{
				int x=(int) sh.getRow(j).getCell(i) .getNumericCellValue();//row wise in each column
			    if(colmax<x)
			    {
			    	colmax=x;
			    }
        	}
        	if(i==0)
        	{
        		sh.createRow(nour).createCell(i).setCellValue(colmax);
        	}
        	else
        	{
        		sh.getRow(nour).createCell(i).setCellValue(colmax);
        	}
        	sh.autoSizeColumn(i);
        }
        //Take write permission on that file
        FileOutputStream fo=new FileOutputStream(f);
		wb.write(fo);
		wb.close();
		fi.close();
		fo.close();
        	
	}

}
