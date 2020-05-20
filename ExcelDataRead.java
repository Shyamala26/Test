//create a method and pass the row number and column no that 
//method will read the data
//of the particular cell
package com.ms.excel;

import java.io.File;
import java.io.IOException;
import java.util.Scanner;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class ExcelDataRead {
	public void readCellData(int reqRow, int reqColumn) throws BiffException, IOException {
		File f = new File("C:\\Users\\Chiranjeevi&Shyamala\\Desktop\\Shyamalaproject.xls");
		Workbook workbook = Workbook.getWorkbook(f);
		Sheet sheet = workbook.getSheet(0);
		int rows = sheet.getRows();
		int columns = sheet.getColumns();
		System.out.println("Total Rows:" + rows);
		System.out.println("Total Columns" + columns);
		for (int rowIndex = 0; rowIndex < rows; rowIndex++) {
			for (int columnIndex = 0; columnIndex < columns; columnIndex++) {
				Cell cell = sheet.getCell(columnIndex, rowIndex);
				System.out.print(cell.getContents());
				System.out.print("\t");
			}
			System.out.println();
		}

		if (reqRow < rows && reqColumn < columns) {
			Cell cell = sheet.getCell(reqColumn, reqRow);
			System.out.println("===> Required Cell Data :" + cell.getContents());
		} else {
			System.out.println("Invalid Row and Column indexs!");
		}

	}

	public static void main(String[] args) throws BiffException, IOException {
		System.out.println("*** Program 1 *** ");
		Scanner s = new Scanner(System.in);
		System.out.print("Enter the Row index to be read : ");
		int reqRow = s.nextInt();
		System.out.print("Enter the Column index to be read : ");
		int reqColumn = s.nextInt();
		ExcelDataRead e = new ExcelDataRead();
		e.readCellData(reqRow, reqColumn);
		s.close();
	}

}
