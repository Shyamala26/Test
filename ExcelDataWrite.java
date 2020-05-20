//create a method and pass the row number and column no that 
//method will read the data
//of the particular cell
package com.ms.excel;
import java.io.File;
import java.io.IOException;
import java.util.Scanner;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
public class ExcelDataWrite {
	public void writeData(int rowCount, int columnCount) throws BiffException, IOException, WriteException {
		if (rowCount > 0 && columnCount > 0) {
			File f = new File("C:\\Users\\Chiranjeevi&Shyamala\\Desktop\\Shyamalaprojectwrite.xls");
			WritableWorkbook workbook = Workbook.createWorkbook(f);
			WritableSheet sheet = workbook.createSheet("shyamala", 3);

			Scanner scanner = new Scanner(System.in);

			for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
				for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
					System.out.println("Enter the Cell data of " + rowIndex + "-" + columnIndex + " :");
					String cellData = scanner.nextLine();
					Label label = new Label(columnIndex, rowIndex, cellData);
					sheet.addCell(label);
				}
			}

			scanner.close();
			workbook.write();
			workbook.close();
			System.out.println("Write Sucessfull!");

		} else {
			System.out.println("Invalid Inputs!");
		}

	}
	public static void main(String[] args) throws BiffException, IOException, WriteException {
		System.out.println("*** Program 1 *** ");
		Scanner s = new Scanner(System.in);
		System.out.print("Enter the Row count : ");
		int rowCount = s.nextInt();
		System.out.print("Enter the Column count : ");
		int columnCount = s.nextInt();
		ExcelDataWrite e = new ExcelDataWrite();
		e.writeData(rowCount, columnCount);
		s.close();
	}

}
