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

public class ReadExcelDataBetweenRowRange {
	public void readCellData(int initialRow, int endRow) throws BiffException, IOException {
		File f = new File("C:\\Users\\Chiranjeevi&Shyamala\\Desktop\\Shyamalaproject.xls");
		Workbook workbook = Workbook.getWorkbook(f);
		Sheet sheet = workbook.getSheet(0);
		int rows = sheet.getRows();
		int columns = sheet.getColumns();
		System.out.println("Total Rows:" + rows);
		System.out.println("Total Columns" + columns);

		if (initialRow <= endRow && endRow < rows) {
			System.out
					.println("===> Required Data of the Row Range Starting from " + initialRow + " to " + endRow + ".");

			// 1-Way - Using getCell() method
			
			for (int rowIndex = initialRow; rowIndex <= endRow; rowIndex++) {
				for (int columnIndex = 0; columnIndex < columns; columnIndex++) {
					Cell cell = sheet.getCell(columnIndex, rowIndex);
					System.out.print(cell.getContents());
					System.out.print("\t");
				}
				System.out.println();
			}


			// 2-Way - Using getRow() method
						for (int rowIndex = initialRow; rowIndex <= endRow; rowIndex++) {
							Cell[] cells = sheet.getRow(rowIndex);
							for (Cell cell : cells) {
								System.out.print(cell.getContents() + "\t");
							}
							System.out.println();
						}
			
		} else {
			System.out.println("Invalid Row range indexs!");
		}
	}

	public static void main(String[] args) throws BiffException, IOException {
		System.out.println("*** Program 1 *** ");
		Scanner s = new Scanner(System.in);
		System.out.print("Enter the Initial Row Index :");
		int initialRow = s.nextInt();
		System.out.print("Enter the End Row Index :");
		int endRow = s.nextInt();
		ReadExcelDataBetweenRowRange e = new ReadExcelDataBetweenRowRange();
		e.readCellData(initialRow, endRow);
		s.close();
	}

}
