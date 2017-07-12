package com.helloselenium.java;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteDataToExcel {

	@SuppressWarnings("resource")
	public static void main(String[] args) {
		try {

			FileInputStream file = new FileInputStream(new File("resources", "actual_excel.xlsx"));
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			XSSFSheet sheet = workbook.getSheetAt(0);

			String line = null;
			int num = 0;
			List<String> lines = new ArrayList<String>();
			FileReader fileReader = new FileReader(new File("resources", "extracted_pdf_data.txt"));
			BufferedReader bufferedReader = new BufferedReader(fileReader);
			while ((line = bufferedReader.readLine()) != null) {
				lines.add(line);
			}

			for (int i = -1; i <= lines.size(); i += 2) {
				if (num == 0) {
					XSSFCell col = sheet.getRow(1).createCell(num);
					col.setCellValue("1");
				} else {
					XSSFCell col = sheet.getRow(1).createCell(num);
					col.setCellValue(lines.get(i));
				}
				num++;
			}

			file.close();

			FileOutputStream outFile = new FileOutputStream(new File("resources", "actual_excel.xlsx"));
			workbook.write(outFile);
			outFile.close();

		} catch (FileNotFoundException fnfe) {
			fnfe.printStackTrace();
		} catch (IOException ioe) {
			ioe.printStackTrace();
		}
	}
}
