package practice;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Sample2 {

	public static void main(String[] args) throws Exception
	{
		//create a new file(.xlsx)
		XSSFWorkbook wb=new XSSFWorkbook();
		Sheet sh=wb.createSheet();
		Row r=sh.createRow(0);
		Cell c=r.createCell(0);
		c.setCellValue("abdul kalam");
		//save in HDD
		sh.autoSizeColumn(0);
		File f=new File("target\\dummy2.xlsx");
		FileOutputStream fo=new FileOutputStream(f);
		wb.write(fo);
		wb.close();
		fo.close();
	
		
	}

}
