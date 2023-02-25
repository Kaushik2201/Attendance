import java.io.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import java.util.*;
import java.util.Calendar;
public class MyAttendence{
	public static void main(String[] args)
	throws FileNotFoundException, IOException
	{
		int cy110=42;
		int cy111=42;
		int cy110a=10;
		int cy111a=3;
		int wo110=42;
		int ma110=10;
		int wo110a=10;
		int ma110a=10;
		int psc=56;
		int psca=14;
		int cs100=42;
		int cs100a=10;
		int cs101=42;
		int cs101a=3;
		int mx = 0;
		Scanner sc = new Scanner(System.in);
		Calendar calendar = Calendar.getInstance();
		Workbook wb = new HSSFWorkbook();
		OutputStream os = new FileOutputStream("Attendence.csv");
		Sheet sheet = wb.createSheet(" Attendence ");
		Row row = sheet.createRow(1);
		Cell cell = row.createCell(1);
		cell.setCellValue("Welcome to your Attendance");
		sheet.addMergedRegion(new CellRangeAddress(1,3,1,3));  
		int rowIndex = cell.getRowIndex();
		int columnIndex = cell.getColumnIndex();
		row = sheet.createRow(5);
		cell = row.createCell(5);
		cell.setCellValue("Classes");
		sheet.addMergedRegion(new CellRangeAddress(5,6,5,6));  
		rowIndex = cell.getRowIndex();
		columnIndex = cell.getColumnIndex();
		row = sheet.createRow(6);
		cell = row.createCell(9);
		cell.setCellValue("Classes Remaining to miss");
		sheet.addMergedRegion(new CellRangeAddress(6,6,9,11));  
		rowIndex = cell.getRowIndex();
		columnIndex = cell.getColumnIndex();
		row = sheet.createRow(9);
		cell = row.createCell(5);
		cell.setCellValue("Chemistry");
		sheet.addMergedRegion(new CellRangeAddress(9,10,5,6));  
		rowIndex = cell.getRowIndex();
		columnIndex = cell.getColumnIndex();
		row = sheet.createRow(12);
		cell = row.createCell(5);
		cell.setCellValue("Mechanics");
		sheet.addMergedRegion(new CellRangeAddress(12,13,5,6));  
		rowIndex = cell.getRowIndex();
		columnIndex = cell.getColumnIndex();
		row = sheet.createRow(15);
		cell = row.createCell(5);
		cell.setCellValue("Maths");
		sheet.addMergedRegion(new CellRangeAddress(15,16,5,6));  
		rowIndex = cell.getRowIndex();
		columnIndex = cell.getColumnIndex();
		row = sheet.createRow(18);
		cell = row.createCell(5);
		cell.setCellValue("Computer ");
		sheet.addMergedRegion(new CellRangeAddress(18,19,5,6));  
		rowIndex = cell.getRowIndex();
		columnIndex = cell.getColumnIndex();
		row = sheet.createRow(21);
		cell = row.createCell(5);
		cell.setCellValue("PSC");
		sheet.addMergedRegion(new CellRangeAddress(21,22,5,6));  
		rowIndex = cell.getRowIndex();
		columnIndex = cell.getColumnIndex();
		row = sheet.createRow(24);
		cell = row.createCell(5);
		cell.setCellValue("Chemistry Lab");
		sheet.addMergedRegion(new CellRangeAddress(24,25,5,6));  
		rowIndex = cell.getRowIndex();
		columnIndex = cell.getColumnIndex();
		row = sheet.createRow(27);
		cell = row.createCell(5);
		cell.setCellValue("Python Lab");
		sheet.addMergedRegion(new CellRangeAddress(27,28,5,6));  
		rowIndex = cell.getRowIndex();
		columnIndex = cell.getColumnIndex();
		System.out.println("Did you miss any classes before ?\n1.Yes\n2.No");
		int miss = sc.nextInt();
		switch (miss) {
			case 1:{
				System.out.println("How many classes of CY110");
				int x = sc.nextInt();
				cy110a-=x;
				x=0;
				System.out.println("How many classes of CY111 (Chemistry Lab) ");
				 x = sc.nextInt();
				cy111a-=x;
				x=0;
				System.out.println("How many classes of WO110");
				x = sc.nextInt();
				wo110a-=x;
				x=0;
				System.out.println("How many classes of MA110");
				x = sc.nextInt();
				ma110a-=x;
				x=0;
				System.out.println("How many classes of PSC");
				 x = sc.nextInt();
				psca-=x;
				x=0;
				System.out.println("How many classes of CS100");
				 x = sc.nextInt();
				cs100a-=x;
				x=0;
				System.out.println("How many classes of CS101  (Python Lab) ");
				 x = sc.nextInt();
				cs101a-=x;
				x=0;

			}
				
				break;
		case 2:{
			System.out.println("You are a good boy");
		}
		break;
		
		default:{
			System.out.println("Please enter the correct number");
		}
				break;
		}
		System.out.println("Day Date and Time :"+calendar.getTime().toString());
		int day = calendar.get(Calendar.DAY_OF_WEEK);
		row = sheet.createRow(3);
		cell = row.createCell(15);
		cell.setCellValue(calendar.getTime().toString());
		sheet.addMergedRegion(new CellRangeAddress(3,3,15,18));  
		 rowIndex = cell.getRowIndex();
		columnIndex = cell.getColumnIndex();
		System.out.println("Did You go to class today ?\n1.Yes\n2.No");
		int go = sc.nextInt();
		switch (go) {
			case 1:{
				System.out.println("Keep Going You are on right track");
				row = sheet.createRow(10);
				cell = row.createCell(9);
				cell.setCellValue(cy110a);
				rowIndex = cell.getRowIndex();
				columnIndex = cell.getColumnIndex();
				row = sheet.createRow(19);
			   cell = row.createCell(9);
			   cell.setCellValue(cs100a);
			   rowIndex = cell.getRowIndex();
			   columnIndex = cell.getColumnIndex();
			   row = sheet.createRow(13);
			   cell = row.createCell(9);
			   cell.setCellValue(wo110a);
			   rowIndex = cell.getRowIndex();
			   columnIndex = cell.getColumnIndex();
			   row = sheet.createRow(22);
				cell = row.createCell(9);
				cell.setCellValue(psca);
				rowIndex = cell.getRowIndex();
				columnIndex = cell.getColumnIndex();
				row = sheet.createRow(16);
				cell = row.createCell(9);
				cell.setCellValue(ma110a);
				rowIndex = cell.getRowIndex();
				columnIndex = cell.getColumnIndex();
				row = sheet.createRow(28);
				cell = row.createCell(9);
				cell.setCellValue(cs101a);
				rowIndex = cell.getRowIndex();
				columnIndex = cell.getColumnIndex();
				row = sheet.createRow(25);
				cell = row.createCell(9);
				cell.setCellValue(cy111a);
				rowIndex = cell.getRowIndex();
				columnIndex = cell.getColumnIndex();
			}
			break;
			case 2:{
			switch (day) {
				case 1:{
					System.out.println("Its Sunday Enjoy your Holiday !");
						row = sheet.createRow(10);
						 cell = row.createCell(9);
						 cell.setCellValue(cy110a);
						 rowIndex = cell.getRowIndex();
						 columnIndex = cell.getColumnIndex();
						 row = sheet.createRow(19);
						cell = row.createCell(9);
						cell.setCellValue(cs100a);
						rowIndex = cell.getRowIndex();
						columnIndex = cell.getColumnIndex();
						row = sheet.createRow(13);
						cell = row.createCell(9);
						cell.setCellValue(wo110a);
						rowIndex = cell.getRowIndex();
						columnIndex = cell.getColumnIndex();
						row = sheet.createRow(22);
						 cell = row.createCell(9);
						 cell.setCellValue(psca);
						 rowIndex = cell.getRowIndex();
						 columnIndex = cell.getColumnIndex();
						 row = sheet.createRow(16);
						 cell = row.createCell(9);
						 cell.setCellValue(ma110a);
						 rowIndex = cell.getRowIndex();
						 columnIndex = cell.getColumnIndex();
						 row = sheet.createRow(28);
						 cell = row.createCell(9);
						 cell.setCellValue(cs101a);
						 rowIndex = cell.getRowIndex();
						 columnIndex = cell.getColumnIndex();
						 row = sheet.createRow(25);
						 cell = row.createCell(9);
						 cell.setCellValue(cy111a);
						 rowIndex = cell.getRowIndex();
						 columnIndex = cell.getColumnIndex();
				}
					
					break;
				case 2:{
					System.out.println("How many classes you missed today ?");
					int m = sc.nextInt();
					for(int i = 0; i<1;i++){
						System.out.println("How many classes of CS100");
						mx = sc.nextInt();
						 cs100a -=mx;
						 mx=0;
						row = sheet.createRow(19);
						cell = row.createCell(9);
						cell.setCellValue(cs100a);
						rowIndex = cell.getRowIndex();
						columnIndex = cell.getColumnIndex();
						 System.out.println("How many classes of WO110");
						 mx = sc.nextInt();
						 wo110a -=mx;
						 mx=0;
						 row = sheet.createRow(13);
						 cell = row.createCell(9);
						 cell.setCellValue(wo110a);
						 rowIndex = cell.getRowIndex();
						 columnIndex = cell.getColumnIndex();
						 System.out.println("How many classes of PSC");
						 mx = sc.nextInt();
						 psca -=mx;
						 mx=0;
						 row = sheet.createRow(22);
						 cell = row.createCell(9);
						 cell.setCellValue(psca);
						 rowIndex = cell.getRowIndex();
						 columnIndex = cell.getColumnIndex();
						 row = sheet.createRow(25);
						 cell = row.createCell(9);
						 cell.setCellValue(cy111a);
						 rowIndex = cell.getRowIndex();
						 columnIndex = cell.getColumnIndex();
					}
				}
				break;
				case 3:{
					System.out.println("How many classes you missed today ?");
					int m = sc.nextInt();
					for(int i = 0; i<1;i++){
						System.out.println("How many classes of PSC");
						mx = sc.nextInt();
						 psca -=mx;
						 mx=0;
						 row = sheet.createRow(22);
						 cell = row.createCell(9);
						 cell.setCellValue(psca);
						 rowIndex = cell.getRowIndex();
						 columnIndex = cell.getColumnIndex();
						 System.out.println("How many classes of CS100");
						 mx = sc.nextInt();
						 cs100a -=mx;
						 mx=0;
						 row = sheet.createRow(19);
						cell = row.createCell(9);
						cell.setCellValue(cs100a);
						rowIndex = cell.getRowIndex();
						columnIndex = cell.getColumnIndex();
						 System.out.println("How many classes of CY110");
						 mx = sc.nextInt();
						 cy110a -=mx;
						 mx=0;
						 row = sheet.createRow(10);
						 cell = row.createCell(9);
						 cell.setCellValue(cy110a);
						 rowIndex = cell.getRowIndex();
						 columnIndex = cell.getColumnIndex();
						 System.out.println("How many classes of MA110");
						 mx=sc.nextInt();
						 ma110a -=mx;
						 mx=0;
						 row = sheet.createRow(16);
						 cell = row.createCell(9);
						 cell.setCellValue(ma110a);
						 rowIndex = cell.getRowIndex();
						 columnIndex = cell.getColumnIndex();
						 System.out.println("Python lab");
						 mx = sc.nextInt();
						 cs100a-=mx;
						 mx=0;
						 row = sheet.createRow(28);
						 cell = row.createCell(9);
						 cell.setCellValue(cs101a);
						 rowIndex = cell.getRowIndex();
						 columnIndex = cell.getColumnIndex();
					}
				}
				break;
				case 4:{
					System.out.println("How many classes you missed today ?");
					int m = sc.nextInt();
					for(int i = 0; i<1;i++){
						System.out.println("How many classes of CY110");
						mx = sc.nextInt();
						 cy110a -=mx;
						 mx=0;
						 row = sheet.createRow(10);
						 cell = row.createCell(9);
						 cell.setCellValue(cy110a);
						 rowIndex = cell.getRowIndex();
						 columnIndex = cell.getColumnIndex();
						 System.out.println("How many classes of MA110");
						 mx = sc.nextInt();
						 ma110a -=mx;
						 mx=0;
						 row = sheet.createRow(16);
						 cell = row.createCell(9);
						 cell.setCellValue(ma110a);
						 rowIndex = cell.getRowIndex();
						 columnIndex = cell.getColumnIndex();
						 System.out.println("How many classes of WO110");
						 mx = sc.nextInt();
						 wo110a -=mx;
						 mx=0;
						 row = sheet.createRow(13);
						 cell = row.createCell(9);
						 cell.setCellValue(wo110a);
						 rowIndex = cell.getRowIndex();
						 columnIndex = cell.getColumnIndex();
						 System.out.println("Chemistry lab");
						 mx = sc.nextInt();
						 cy111a-=mx;
						 mx=0;
						 row = sheet.createRow(25);
						 cell = row.createCell(9);
						 cell.setCellValue(cy111a);
						 rowIndex = cell.getRowIndex();
						 columnIndex = cell.getColumnIndex();
						 System.out.println("How many classes of PSC");
						 mx = sc.nextInt();
						  psca -=mx;
						  mx=0;
						  row = sheet.createRow(22);
						  cell = row.createCell(9);
						  cell.setCellValue(psca);
						  rowIndex = cell.getRowIndex();
						  columnIndex = cell.getColumnIndex();
					}
				}
				break;
				case 5:{
					System.out.println("How many classes you misses today");
					int m = sc.nextInt();
					for(int i = 0; i<1;i++){
						System.out.println("How many classes of CY110");
						mx = sc.nextInt();
						 cy110a -=mx;
						 mx=0;
						 row = sheet.createRow(10);
						 cell = row.createCell(9);
						 cell.setCellValue(cy110a);
						 rowIndex = cell.getRowIndex();
						 columnIndex = cell.getColumnIndex();
						 System.out.println("How many classes of WO110");
						 mx = sc.nextInt();
						 wo110a -=mx;
						 mx=0;
						 row = sheet.createRow(13);
						 cell = row.createCell(9);
						 cell.setCellValue(wo110a);
						 rowIndex = cell.getRowIndex();
						 columnIndex = cell.getColumnIndex();
						 System.out.println("How many classes of MA110");
						 mx = sc.nextInt();
						 ma110a -=mx;
						 mx=0;
						 row = sheet.createRow(16);
						 cell = row.createCell(9);
						 cell.setCellValue(ma110a);
						 rowIndex = cell.getRowIndex();
						 columnIndex = cell.getColumnIndex();
						 System.out.println("CS100");
						 mx = sc.nextInt();
						 cs100a-=mx;
						 mx=0;
						 row = sheet.createRow(19);
						cell = row.createCell(9);
						cell.setCellValue(cs100a);
						rowIndex = cell.getRowIndex();
						columnIndex = cell.getColumnIndex();
						 System.out.println("How many classes of PSC");
						 mx = sc.nextInt();
						  psca -=mx;
						  mx=0;
						  row = sheet.createRow(22);
						  cell = row.createCell(9);
						  cell.setCellValue(psca);
						  rowIndex = cell.getRowIndex();
						  columnIndex = cell.getColumnIndex();
					}
				}
				break;
				case 6:{
					System.out.println("Its Friday Enjoy your Holiday !");
					row = sheet.createRow(10);
					cell = row.createCell(9);
					cell.setCellValue(cy110a);
					rowIndex = cell.getRowIndex();
					columnIndex = cell.getColumnIndex();
					row = sheet.createRow(19);
				   cell = row.createCell(9);
				   cell.setCellValue(cs100a);
				   rowIndex = cell.getRowIndex();
				   columnIndex = cell.getColumnIndex();
				   row = sheet.createRow(13);
				   cell = row.createCell(9);
				   cell.setCellValue(wo110a);
				   rowIndex = cell.getRowIndex();
				   columnIndex = cell.getColumnIndex();
				   row = sheet.createRow(22);
					cell = row.createCell(9);
					cell.setCellValue(psca);
					rowIndex = cell.getRowIndex();
					columnIndex = cell.getColumnIndex();
					row = sheet.createRow(16);
					cell = row.createCell(9);
					cell.setCellValue(ma110a);
					rowIndex = cell.getRowIndex();
					columnIndex = cell.getColumnIndex();
					row = sheet.createRow(28);
					cell = row.createCell(9);
					cell.setCellValue(cs101a);
					rowIndex = cell.getRowIndex();
					columnIndex = cell.getColumnIndex();
					row = sheet.createRow(25);
					cell = row.createCell(9);
					cell.setCellValue(cy111a);
					rowIndex = cell.getRowIndex();
					columnIndex = cell.getColumnIndex();
				}
				break;
				case 7:{
					System.out.println("Its Saturday Enjoy your Holiday !");
					row = sheet.createRow(10);
					cell = row.createCell(9);
					cell.setCellValue(cy110a);
					rowIndex = cell.getRowIndex();
					columnIndex = cell.getColumnIndex();
					row = sheet.createRow(19);
				   cell = row.createCell(9);
				   cell.setCellValue(cs100a);
				   rowIndex = cell.getRowIndex();
				   columnIndex = cell.getColumnIndex();
				   row = sheet.createRow(13);
				   cell = row.createCell(9);
				   cell.setCellValue(wo110a);
				   rowIndex = cell.getRowIndex();
				   columnIndex = cell.getColumnIndex();
				   row = sheet.createRow(22);
					cell = row.createCell(9);
					cell.setCellValue(psca);
					rowIndex = cell.getRowIndex();
					columnIndex = cell.getColumnIndex();
					row = sheet.createRow(16);
					cell = row.createCell(9);
					cell.setCellValue(ma110a);
					rowIndex = cell.getRowIndex();
					columnIndex = cell.getColumnIndex();
					row = sheet.createRow(28);
					cell = row.createCell(9);
					cell.setCellValue(cs101a);
					rowIndex = cell.getRowIndex();
					columnIndex = cell.getColumnIndex();
					row = sheet.createRow(25);
					cell = row.createCell(9);
					cell.setCellValue(cy111a);
					rowIndex = cell.getRowIndex();
					columnIndex = cell.getColumnIndex();
				}
				break;

				default: {
					System.out.println("Invalid Date !");
				}
					break;
			}
			
		}
				break;

	}
	wb.write(os);
	wb.close();
	sc.close();
}
}
