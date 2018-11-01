package excelget;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Excelee {

	public static void main(String[] args) {

		try {

			// 첫번째 문서 읽기
			FileInputStream fis = new FileInputStream("/Users/a503-02/Documents/exceltestfile/GGD_RouteInfo_M.xls");
			HSSFWorkbook workbook = new HSSFWorkbook(fis);
			HSSFSheet sheet = workbook.getSheetAt(0); // 시트 수
			int rows = sheet.getPhysicalNumberOfRows(); // 해당시트의 행의 총 수

			String[] city = new String[rows];
			String[] busno = new String[rows];
			String[] begin = new String[rows];

			// 두번째 문서 읽기
			FileInputStream fis2 = new FileInputStream("/Users/a503-02/Documents/exceltestfile/bustotal.xls");
			HSSFWorkbook workbook2 = new HSSFWorkbook(fis2);

			HSSFSheet[] sheets = new HSSFSheet[3];
			int[] rowrow = new int[3];

			for (int i = 0; i < 3; i++) {
				sheets[i] = workbook2.getSheetAt(i); // 각 시트
				rowrow[i] = sheets[i].getPhysicalNumberOfRows(); // 각 시트의 행 수
			}

			// 새로운 엑셀 생성
			HSSFWorkbook workbook1 = new HSSFWorkbook(); // 새 엑셀 생성
			HSSFSheet[] makesheets = new HSSFSheet[3]; // 세개의 시트 생성

			
			makesheets[0] = workbook1.createSheet("first");
			makesheets[1] = workbook1.createSheet("second");
			makesheets[2] = workbook1.createSheet("third");
			
			int count=0;

			// 첫번째 엑셀 파일 읽기

			for (int rowindex = 1; rowindex < rows; rowindex++) {
				HSSFRow row = sheet.getRow(rowindex); // 첫번째 엑셀파일의 행을 읽어온다.

				HSSFCell cell1 = row.getCell(1); // 첫번째 파일의 셀에 담겨있는 값을 읽는다.
				HSSFCell cell3 = row.getCell(3);
				HSSFCell cell4 = row.getCell(4);

				String value1 = cell1.getStringCellValue().trim(); // 값을 string 형태로 바꾼다.
				String value3 = cell3.getStringCellValue().trim();
				String value4 = cell4.getStringCellValue().trim();

				city[rowindex-1] = value1;
				busno[rowindex-1] = value3;
				begin[rowindex-1] = value4;
				
				//count++;

				//System.out.println(value1 + "   " + value3 + "   " + value4);

			}

			
			
			
			
			System.out.println("==================두번째 파일===================");

			// 두번째 엑셀 파일 읽기
			String cityname1 = new String();
			for (int sheetnum = 0; sheetnum < 3; sheetnum++) { // 두번째 엑셀파일

				
				for (int k = 1; k < rowrow[sheetnum]; k++) {

					HSSFRow row2 = sheets[sheetnum].getRow(k); // 두번째 엑셀파일의 행을 읽어온다.

					HSSFCell cell11 = row2.getCell(1); // 두번째 파일의 셀에 담겨있는 값을 읽는다.
					HSSFCell cell22 = row2.getCell(2);
					HSSFCell cell33 = row2.getCell(3);
					HSSFCell cell44 = row2.getCell(4);

					String value11 = cell11.getStringCellValue().trim();// 값을 string 형태로 바꾼다.
					String value22 = cell22.getStringCellValue().trim();
					String value33 = cell33.getStringCellValue().trim();
					String value44 = cell44.getStringCellValue().trim();

					//System.out.println(value11 + "   " + value22 + "   " + value33 + "   " + value44);

					
					HSSFRow rrow1 = makesheets[sheetnum].createRow(k); // 해당 행 생성

					HSSFCell cel1 = rrow1.createCell(0); // 생성된 엑셀파일의 열 생성
					cel1.setCellValue(value11);
					HSSFCell cel2 = rrow1.createCell(1);
					cel2.setCellValue(value22);
					HSSFCell cel3 = rrow1.createCell(2);
					cel3.setCellValue(value33);
					HSSFCell cel4 = rrow1.createCell(3);
					cel4.setCellValue(value44);

				
					
					for (int i = 0; i < city.length-1; i++) {

						if (busno[i].equals(value22) && begin[i].equals(value44)) {
							HSSFCell cel6 = rrow1.createCell(5);
							cel6.setCellValue(city[i]);
							
							cityname1=city[i];
						
							
						}

					}
					
					
					
					HSSFCell cel6 = rrow1.createCell(5);
					cel6.setCellValue(cityname1);
					
					
					
					
					
					
				}
				
			}



			FileOutputStream fileoutputstream = new FileOutputStream("/Users/a503-02/Documents/exceltestfile/totalbuslocationex2.xls");
			workbook1.write(fileoutputstream);
			fileoutputstream.close();
			System.out.println("엑셀파일생성성공");
			


		} catch (Exception e) {
			System.out.println(e.getMessage());
		}

	}

}
