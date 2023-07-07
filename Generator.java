package main;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.BufferedWriter;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Iterator;

public class Generator {
	public static void main(String[] args) {
		excelToBeanClass();
	}

	static String excelToBeanClass() {
		String excelFilePath = "C:/Users/HarishRathinamS/Downloads/Logistics/Domestic/POJOImport.xlsx"; // SpreadSheet File Path
		String pojoClassContent = null;
		try (FileInputStream fileInputStream = new FileInputStream(excelFilePath);
				Workbook workbook = new XSSFWorkbook(fileInputStream)) {

			String sheetName = null;
			int sheetCount = workbook.getNumberOfSheets(); // Get number of sheets in Excel File
			String packageNameString = "package com.wattsavvy.core.datamodel;";
			
			for (int index = 0; index < sheetCount; index++) {
				Sheet sheetIndex = workbook.getSheetAt(index); // Iterate sheet one by one
				sheetName = sheetIndex.getSheetName();
				System.out.println("Sheet Name: " + sheetName);
				Sheet sheet = workbook.getSheet(sheetName);
				Iterator<Row> rowIterator = sheet.iterator();// Iterates row by row
				rowIterator.next(); // Skip the header row
				StringBuilder pojoClassBuilder = new StringBuilder();
				pojoClassBuilder.append(packageNameString+"\n\n");
				pojoClassBuilder.append("public class " + sheetName + " {\n"); // Defines the Class Name which is same as Sheet Name
				while (rowIterator.hasNext()) {
					Row row = rowIterator.next();

					Cell dataTypeCell = row.getCell(0);
					Cell variableNameCell = row.getCell(1);

					String dataType = dataTypeCell.getStringCellValue();
					String variableName = variableNameCell.getStringCellValue();
					pojoClassBuilder.append("\tprivate ").append(dataType).append(" ").append(variableName)
							.append(";\n");
				}
				pojoClassBuilder.append("}");
				pojoClassContent = pojoClassBuilder.toString();
				writeNewFile(sheetName,pojoClassContent); //Sheet Name is file name , Generated POJO Class
			}

		} catch (IOException e) {
			e.printStackTrace();
		}
		return pojoClassContent;
	}
	
	public static void writeNewFile(String className, String classContent) {
		String outputPath = "C:/Users/HarishRathinamS/Downloads/Logistics/Domestic/"+className+".java";
		try (BufferedWriter writer = new BufferedWriter(new FileWriter(outputPath))) {
			writer.write(classContent); //Pojo is written with the Sheet Name as Class Name
			System.out.println("Java class file created successfully.");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}