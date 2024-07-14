package vn.edu.likelion.baitap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Scanner;

public class BaitapApplication {
	public static void main(String[] args) {
		// Tạo một mảng tạm để lưu trữ dữ liệu từ fileA
		ArrayList<String> arrStringTemp = new ArrayList<>();

		// ======================================================
		// Lấy ngày hiện tại và định dạng lại kiểu (ddMMyyyy)
		// ======================================================
		LocalDate currentDateTime = LocalDate.now();
		DateTimeFormatter formatter = DateTimeFormatter.ofPattern("ddMMyyyy");
		String formattedDate = currentDateTime.format(formatter);

		System.out.print("Vui lòng nhập vào tên file: ");
		Scanner scan = new Scanner(System.in);
		String filename = scan.nextLine() + "_" + formattedDate + ".xlsx";
		 System.out.println(filename);

		// Tạo đối tượng BufferedWriter
		BufferedReader fileReader = null;

		try {
			// Đọc file bằng BufferedWriter
			fileReader = new BufferedReader(new FileReader("StudentsList.txt"));

			// Lưu tất cả dữ liệu khi đọc từng dòng code vào arrStringTemp
			String content;
			String studentName;
			while ((content = fileReader.readLine()) != null) {
				studentName  = content.split("	")[1];	// Cắt chuỗi thành mảng, lấy phần tử thứ 1 là họ và tên
				arrStringTemp.add(studentName);
			}
			System.out.println("Đã sao chép vào bộ nhố tạm.");
//			System.out.println(arrStringTemp);



			// Lấy file excel vật lý
			FileInputStream fis = new FileInputStream(new File("output.xlsx"));
			// Create Workbook
			Workbook workbook = new XSSFWorkbook(fis);
			// Lấy sheet đầu tiên từ workbook
			Sheet sheet = workbook.getSheetAt(0);


			int dataIndex = 4;
			for (int i = 0; i < arrStringTemp.size(); i++) {
				Row row = sheet.getRow(dataIndex);
				if (row == null) {
					row = sheet.createRow(dataIndex);
				}
				Cell cell = row.getCell(1); // Cột B (duy nhất 1 cột B)
				if (cell == null) {
					cell = row.createCell(1);
				}

				cell.setCellValue(arrStringTemp.get(i));
				dataIndex++;
			}

			FileOutputStream fos = new FileOutputStream("output.xlsx");
			workbook.write(fos);
			workbook.close();
			fos.close();

			System.out.println("Đã ghi dữ liệu vào file Excel.");


//			Thread writeFileThread = new WriteFileThread(filename, arrStringTemp);
//			writeFileThread.start();






		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			throw new RuntimeException(e);
		} finally {
			try {
				if (fileReader != null) {
					fileReader.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}


	public static void ReadExcelFile( String filePath){
        try {
			// Lấy file excel vật lý
			FileInputStream fis = new FileInputStream(new File(filePath));

			//Tạo ra 1 cái workBook từ file vật lý
			Workbook workbook = new XSSFWorkbook(fis);

			// Lấy sheet đầu tiên từ workbook
			Sheet sheet = workbook.getSheetAt(0);

			//Duyệt từng row trong sheet
			for (Row row : sheet) {
				// Duyệt từng cell trong row
				for (Cell cell : row) {
					// Kiểm tra xem cell nó đang có kiểu là gì ?
					switch (cell.getCellType()) {
						case STRING:
							System.out.println(cell.getStringCellValue());
							break;
						case NUMERIC:
							System.out.println(cell.getNumericCellValue());
							break;
						case BOOLEAN:
							System.out.println(cell.getBooleanCellValue());
							break;
						case FORMULA:
							System.out.println(cell.getCellFormula());
							break;
						case BLANK:
							System.out.println("");
							break;
						case ERROR:
							System.out.println(cell.getErrorCellValue());
							break;

						default:
							System.out.println(cell.getDateCellValue());
							break;
					}
				}
			}

		} catch (IOException io){
			io.printStackTrace();
		}
	}




	public static class WriteFileThread extends Thread {
		private String filename;
		private ArrayList<String> arrStringTemp;

		public WriteFileThread(String filename, ArrayList<String> arrStringTemp) {
			this.filename = filename;
			this.arrStringTemp = arrStringTemp;
		}

		@Override
		public void run() {
			System.out.println(Thread.currentThread().getName() + ": Writing file " + filename + "...");

			try {
				// Lấy file excel vật lý
				FileInputStream fis = new FileInputStream(new File("output.xlsx"));
				// Tạo ra 1 đối tượng xử lý file excel
				XSSFWorkbook workbook = new XSSFWorkbook();

				// Tạo 1 đối tượng xử lý tài liệu docx
				XWPFDocument document = new XWPFDocument();
			} catch (IOException io){
				io.printStackTrace();
			}

		}
	}


















}
