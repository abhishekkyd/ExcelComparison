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

public class CompareExcel {

	@SuppressWarnings({ "resource", "deprecation" })
	public static void main(String[] args) {
		try {

			FileInputStream file = new FileInputStream(new File("resources", "result_excel.xlsx"));
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(0);

			FileInputStream file1 = new FileInputStream(new File("resources", "expected_excel.xlsx"));
			XSSFWorkbook workbook1 = new XSSFWorkbook(file1);
			XSSFSheet sheet1 = workbook1.getSheetAt(0);

			FileInputStream file2 = new FileInputStream(new File("resources", "actual_excel.xlsx"));
			XSSFWorkbook workbook2 = new XSSFWorkbook(file2);
			XSSFSheet sheet2 = workbook2.getSheetAt(0);

			int num = 1;

			for (int i = 0; i <= 5; i++) {
				XSSFCell col = sheet.getRow(1).createCell(i);
				XSSFCell col1 = sheet1.getRow(1).getCell(i);
				String expected = null;
				if (col1.getCellTypeEnum().toString().equalsIgnoreCase("string")) {
					expected = col1.getStringCellValue();
				} else if (col1.getCellTypeEnum().toString().equalsIgnoreCase("numeric")) {
					expected = String.valueOf(col1.getNumericCellValue());
				}
				XSSFCell col2 = sheet2.getRow(1).getCell(i);
				String actual = null;
				if (col2.getCellTypeEnum().toString().equalsIgnoreCase("string")) {
					actual = col2.getStringCellValue();
				} else if (col2.getCellTypeEnum().toString().equalsIgnoreCase("numeric")) {
					actual = String.valueOf(col2.getNumericCellValue());
				}

				if (i == 0) {
					col.setCellValue(num);
				} else {

					if (expected.equals(actual)) {
						col.setCellValue("Pass");
					} else {
						col.setCellValue("Fail");
					}
				}
			}

			file.close();
			file1.close();
			file2.close();

			FileOutputStream outFile = new FileOutputStream(new File("resources", "result_excel.xlsx"));
			workbook.write(outFile);
			outFile.close();

		} catch (FileNotFoundException fnfe) {
			fnfe.printStackTrace();
		} catch (IOException ioe) {
			ioe.printStackTrace();
		}
	}
}
