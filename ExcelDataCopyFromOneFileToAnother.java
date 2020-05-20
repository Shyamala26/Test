//Read Data from file1 and write the data from file2
//method will read the data
//of the particular cell
package com.ms.excel;

import java.io.File;
import java.io.IOException;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;
public class ExcelDataCopyFromOneFileToAnother {
	public void copy() throws BiffException, IOException, RowsExceededException, WriteException {
		File f = new File("C:\\Users\\Chiranjeevi&Shyamala\\Desktop\\Shyamalaproject.xls");
		Workbook readWorkbook = Workbook.getWorkbook(f);
		Sheet readSheet = readWorkbook.getSheet(0);
		int rows = readSheet.getRows();
		int columns = readSheet.getColumns();
		System.out.println("Total Rows:" + rows);
		System.out.println("Total Columns" + columns);

		File fw = new File("C:\\Users\\Chiranjeevi&Shyamala\\Desktop\\ShyamalaAssignment5.xls");
		WritableWorkbook writeWorkbook = Workbook.createWorkbook(fw);
		WritableSheet writeSheet = writeWorkbook.createSheet("shyamala", 3);

		for (int rowIndex = 0; rowIndex < rows; rowIndex++) {
			for (int columnIndex = 0; columnIndex < columns; columnIndex++) {
				Cell cell = readSheet.getCell(columnIndex, rowIndex);
				Label label = new Label(columnIndex, rowIndex, cell.getContents());
				writeSheet.addCell(label);
			}
		}
		writeWorkbook.write();
		writeWorkbook.close();
		System.out.println("Copied Sucessfully!");
	}

	public static void main(String[] args) throws BiffException, IOException, RowsExceededException, WriteException {
		System.out.println("**Program to read data from File1 and Write the data from file2**");
		ExcelDataCopyFromOneFileToAnother e = new ExcelDataCopyFromOneFileToAnother();
		e.copy();

	}

}
