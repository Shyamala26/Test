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
public class ReadExcelDataBasedOnRow {
	public void readCellData(int reqRow) throws BiffException, IOException {
		File f = new File("C:\\Users\\Chiranjeevi&Shyamala\\Desktop\\Shyamalaproject.xls");
		Workbook workbook = Workbook.getWorkbook(f);
		Sheet sheet = workbook.getSheet(0);
		int rows = sheet.getRows();
		int columns = sheet.getColumns();
		System.out.println("Total Rows:" + rows);
		System.out.println("Total Columns" + columns);

		// 2-Way - Using getCell() method
		if (reqRow < rows) {
			System.out.println("===> Required Data of the Row Index :" + reqRow);			
			for (int columnIndex = 0; columnIndex < columns; columnIndex++) {
				Cell cell = sheet.getCell(columnIndex, reqRow);
				System.out.print(cell.getContents() + "\t");
			}

			System.out.println();
			// 2-Way - Using getRow() method
			Cell[] cells = sheet.getRow(reqRow);
			for (Cell cell : cells) {
				System.out.print(cell.getContents() + "\t");
			}

		} else {
			System.out.println("Invalid Row and Column indexs!");
		}
	}

	public static void main(String[] args) throws BiffException, IOException {
		System.out.println("*** Program 1 *** ");
		Scanner s = new Scanner(System.in);
		System.out.print("Enter the Row index to be read : ");
		int reqRow = s.nextInt();
		ReadExcelDataBasedOnRow e = new ReadExcelDataBasedOnRow();
		e.readCellData(reqRow);
		s.close();
	}

}
