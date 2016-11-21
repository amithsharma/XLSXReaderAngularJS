package com.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import javax.ws.rs.GET;
import javax.ws.rs.Path;
import javax.ws.rs.Produces;
import javax.ws.rs.core.MediaType;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@Path("/excelService")
public class ExcelService {

	@GET
	@Path("/excelData")
	@Produces(MediaType.TEXT_PLAIN)
	public String getExcelDataInJSON() {
		String excelFilePath = "C://dev workspace/XLSXReaderAngularJS/src/AngulatTest.xlsx";
		String jsonStart = "[      ";
		String jsonEnd = "    ]";
		String json_1 = "\": \"";
		String json_2 = "\", \"";
		
		StringBuffer jsonOutput = new StringBuffer("");
		try {
			FileInputStream inputStream = new FileInputStream(new File(excelFilePath));

			Workbook workbook = null;

			if (excelFilePath.endsWith("xlsx")) {
				workbook = new XSSFWorkbook(inputStream);
			} else if (excelFilePath.endsWith("xls")) {
				workbook = new HSSFWorkbook(inputStream);
			} else {
				throw new IllegalArgumentException("The specified file is not Excel file");
			}

			Sheet firstSheet = workbook.getSheetAt(0);
			Iterator<Row> iterator = firstSheet.iterator();

			jsonOutput.append(jsonStart);

			Row firstRow = firstSheet.getRow(0);
			int rowCount = firstSheet.getLastRowNum() + 1;
			int columnCount = firstRow.getLastCellNum() + 1;
			// System.out.println(rowCount);
			// System.out.println(columnCount);

			for (int rowNum = 1; rowNum < firstSheet.getLastRowNum() + 1; rowNum++) {
				Row row = firstSheet.getRow(rowNum);
				jsonOutput.append("{\"");
				for (int cellNum = 0; cellNum < columnCount - 1; cellNum++) {
					Cell cell = row.getCell(cellNum);
					jsonOutput.append(firstRow.getCell(cellNum).getStringCellValue());
					jsonOutput.append(json_1);
					if (null != cell) {
						switch (cell.getCellType()) {
						case Cell.CELL_TYPE_STRING:
							jsonOutput.append(cell.getStringCellValue());
							break;
						case Cell.CELL_TYPE_BOOLEAN:
							jsonOutput.append(cell.getBooleanCellValue());
							break;
						case Cell.CELL_TYPE_NUMERIC:
							jsonOutput.append(cell.getNumericCellValue());
							break;
						}
					} else {
						jsonOutput.append("");
					}
					if (cellNum != columnCount - 2) {
						jsonOutput.append(json_2);
					}
				}
				jsonOutput.append("\"}");
				if (rowNum != rowCount - 1) {
					jsonOutput.append(",");
				}
			}
			jsonOutput.append(jsonEnd);
			System.out.println(jsonOutput);

			workbook.close();
			inputStream.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

		return jsonOutput.toString();
	}
}
