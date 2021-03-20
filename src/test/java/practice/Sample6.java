package practice;

import java.io.File;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Sample6 {

	public static void main(String[] args) 
	{
		//open a folder and collect contents(subfolders and files)of that folder
		File f1=new File("c:\\");
		File[] l=f1.listFiles();
		//create a new excel file(.xlsx)
		XSSFWorkbook wwb=new XSSFWorkbook();
		

	}

}
