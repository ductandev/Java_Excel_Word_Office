package vn.edu.likelion.baitap;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Base64;
import java.util.Scanner;

public class Cach2 {
	/**
	 * Workbook: đại diện cho một file Excel. Nó được triển khai dưới hai class là: HSSFWorkbook và XSSFWorkbook tương ứng cho định dạng .xls và .xlsx .
	 * Sheet: đại diện cho một bảng tính Excel (một file Excel có thể có nhiều Sheet). Nó có 2 class là HSSFSheet và XSSFSheet.
	 * Row: đại diện cho một hàng trong một bảng tính (Sheet). Nó có 2 class là HSSFRow và XSSFRow.
	 * Cell: đại diện cho một ô trong một hàng (Row). Tương tự nó cũng có 2 class là HSSFCell and XSSFCell.
	 */

	public static final int COLUMN_INDEX_STT       = 0;
	public static final int COLUMN_INDEX_ID        = 1;
	public static final int COLUMN_INDEX_NAME      = 2;
	public static final int COLUMN_INDEX_STATUS    = 3;
	private static CellStyle cellStyleFormatNumber = null;

	public static void main(String[] args) {
		try {
			LocalDate currentDateTime = LocalDate.now();
			DateTimeFormatter formatter = DateTimeFormatter.ofPattern("ddMMyyyy");
			String formattedDate = currentDateTime.format(formatter);	// Lấy ngày hiện tại và định dạng lại kiểu (ddMMyyyy)

			System.out.print("Vui lòng nhập vào tên file: ");
			Scanner scan = new Scanner(System.in);
			String filePath = scan.nextLine() + "_" + formattedDate;

			final ArrayList<Student> students = getListStudents("StudentsList.txt");	// Lấy danh sách học sinh lưu vào arrListStudents
			writeExcel(students, filePath);

		} catch (IOException e) {
			e.printStackTrace();
		}
	}


	public static void writeExcel(ArrayList<Student> students, String filePath) throws IOException {
		String excelFilePath = filePath + ".xlsx";
		String wordFilePath = filePath + ".docx";

		// Create Workbook
		Workbook workbook = getWorkbook(excelFilePath);

		// Create sheet
		Sheet sheet = workbook.createSheet("Students"); // Create sheet with sheet name

		int rowIndex = 0;

		// Write header
		writeHeader(sheet, rowIndex);

		// Write data
		rowIndex++;
		for (Student studentCurrent : students) {
			if(studentCurrent.getIsActive().equals("0")) {			// lỌC RA danh sách học sinh có mặt
				// Create row
				Row row = sheet.createRow(rowIndex);
				// Write data on row
				writeStudent(studentCurrent, row);
				rowIndex++;

			} else {
				writeWordFile(studentCurrent, wordFilePath);
			}
		}

		// Auto resize column witdth
		int numberOfColumn = sheet.getRow(0).getPhysicalNumberOfCells();
		autosizeColumn(sheet, numberOfColumn);

		// Create file excel
		createOutputFile(workbook, excelFilePath);
		System.out.println("Done!!!");
	}


	// =====================================================
	// 					CREATE WORKBOOK
	// =====================================================
	private static Workbook getWorkbook(String excelFilePath){
        Workbook workbook = null;

        if (excelFilePath.endsWith("xlsx")) {
            workbook = new XSSFWorkbook();
        } else if (excelFilePath.endsWith("xls")) {
            workbook = new HSSFWorkbook();
        } else {
            throw new IllegalArgumentException("The specified file is not Excel file");
        }

        return workbook;
    }

	// =====================================================
	// 				CREATE LIST STUDENT DATA
	// =====================================================
	private static ArrayList<Student> getListStudents(String excelFilePath) throws IOException {
		ArrayList<Student> arrListStudents = new ArrayList<>();				// Tạo một mảng tạm để lưu trữ dữ liệu từ fileA
		BufferedReader fileReader = null;									// Tạo đối tượng BufferedWriter
		fileReader = new BufferedReader(new FileReader(excelFilePath));		// Đọc file bằng BufferedWriter

		String line;
		String[] data;		// ouput: data = [17, Nguyễn Đức Tấn, 0]
		Student student;	// Khởi tạo đối tượng student

		while ((line = fileReader.readLine()) != null) {
			data  = line.split("	");
			data[1] = Base64.getEncoder().encodeToString(data[1].getBytes());	// MÃ HÓA DỮ LIỆU
			student = new Student(data[0], data[1], data[2]);
			arrListStudents.add(student);
		}

		//System.out.println(arrListStudents);
		return arrListStudents;
	}


	// =====================================================
	// 				WRITE HEADER WITH FORMAT
	// =====================================================
	private static void writeHeader(Sheet sheet, int rowIndex) {
		// create CellStyle
		CellStyle cellStyle = createStyleForHeader(sheet);

		// Create row
		Row row = sheet.createRow(rowIndex);

		// Create cells
		Cell cell = row.createCell(COLUMN_INDEX_STT);
		cell.setCellStyle(cellStyle);
		cell.setCellValue("Stt");

		cell = row.createCell(COLUMN_INDEX_ID);
		cell.setCellStyle(cellStyle);
		cell.setCellValue("Id");

		cell = row.createCell(COLUMN_INDEX_NAME);
		cell.setCellStyle(cellStyle);
		cell.setCellValue("Name");

		cell = row.createCell(COLUMN_INDEX_STATUS);
		cell.setCellStyle(cellStyle);
		cell.setCellValue("Status");
	}

	// =====================================================
	// 					    WRITE DATA
	// =====================================================
	private static void writeStudent(Student student, Row row) {
		int currentRow = row.getRowNum();
		Cell cell = row.createCell(COLUMN_INDEX_STT);
		cell.setCellValue(currentRow);

		cell = row.createCell(COLUMN_INDEX_ID);
		cell.setCellValue(student.getId());

		cell = row.createCell(COLUMN_INDEX_NAME);
		cell.setCellValue(student.getName());

		cell = row.createCell(COLUMN_INDEX_STATUS);
		cell.setCellValue(student.getIsActive().equals("0") ? "Có mặt" : "Vắng");
	}

	// =====================================================
	// 				CREATE CELL STYLE FOR HEADER
	// =====================================================
	private static CellStyle createStyleForHeader(Sheet sheet) {
		// Create font
		Font font = sheet.getWorkbook().createFont();
		font.setFontName("Times New Roman");
		font.setBold(true);
		font.setFontHeightInPoints((short) 14); // font size
		font.setColor(IndexedColors.WHITE.getIndex()); // text color

		// Create CellStyle
		CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setFont(font);

		// Center alignment
		cellStyle.setAlignment(HorizontalAlignment.CENTER);
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

		cellStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cellStyle.setBorderBottom(BorderStyle.THIN);
		return cellStyle;
	}

	// =====================================================
	// 					WRITE WORD FILE
	// =====================================================
	private static void writeWordFile(Student student, String wordFilePath) {
		XWPFDocument document;
		File file = new File(wordFilePath);

		// Kiểm tra nếu file tồn tại
		if (file.exists()) {
			try (FileInputStream fis = new FileInputStream(file)) {
				document = new XWPFDocument(fis);
			} catch (IOException e) {
				e.printStackTrace();
				return;
			}
		} else {
			document = new XWPFDocument();
		}

		// Tạo ra 1 đoạn văn bản mới
		XWPFParagraph paragraph = document.createParagraph();

		// Tạo câu văn
		XWPFRun run = paragraph.createRun();
		run.setText(student.getId() + "	" + student.getName() + "	Vắng mặt" + "\n");

		try (FileOutputStream fos = new FileOutputStream(wordFilePath)) {
			// Ghi các giá trị của document vào file word
			document.write(fos);
			System.out.println("Đã tạo hoặc cập nhật file docx thành công !!!");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	// Auto resize column width
	private static void autosizeColumn(Sheet sheet, int lastColumn) {
		for (int columnIndex = 0; columnIndex < lastColumn; columnIndex++) {
			sheet.autoSizeColumn(columnIndex);
		}
	}

	// Create output file
	private static void createOutputFile(Workbook workbook, String excelFilePath) throws IOException {
		try (OutputStream os = new FileOutputStream(excelFilePath)) {
			workbook.write(os);
		}
	}

}
