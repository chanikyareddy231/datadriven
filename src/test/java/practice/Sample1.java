package practice;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class Sample1 {

	public static void main(String[] args) throws Exception
	{
		//create a new file(.xls)
		HSSFWorkbook wb=new HSSFWorkbook();
		Sheet sh=wb.createSheet();
		Row r=sh.createRow(0);
		Cell c=r.createCell(0);
		c.setCellValue("abdul kalam");
		//save in HDD
		sh.autoSizeColumn(0);
		File f=new File("target\\dummy1.xls");
		FileOutputStream fo=new FileOutputStream(f);
		wb.write(fo);
		wb.close();
		fo.close();
	
		
	}

}
