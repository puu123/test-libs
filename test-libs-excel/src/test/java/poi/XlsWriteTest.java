package poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.time.chrono.MinguoEra;
import java.util.Date;
import java.util.function.Consumer;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.WorkbookUtil;



import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import com.github.mygreen.cellformatter.POICellFormatter;

public class XlsWriteTest {

	@Rule
	public TemporaryFolder tempFolder = new TemporaryFolder();

	@Test
	public void test() throws Exception {

		try (FileOutputStream out = new FileOutputStream(new File("foo.xls")); 
				Workbook wb = new HSSFWorkbook();) {
			
			String safeName = WorkbookUtil.createSafeSheetName("['aaa's test*?]");
			Sheet sheet1 = wb.createSheet(safeName);

			CreationHelper createHelper = wb.getCreationHelper();
			//createHelper.createFormulaEvaluator().evaluate(cell)
			short style = createHelper.createDataFormat().getFormat("yyyy-mm-ddThh:mm");
		
			Row row = sheet1.createRow(0);
			Cell cell = row.createCell(0);
			cell.setCellValue(1);
			row.createCell(1).setCellValue(1.2);
			row.createCell(2).setCellValue(createHelper.createRichTextString("sample string"));
			setDate(row.createCell(3), new Date(), wb);

			wb.write(out);
		}
		
	}
	
	@Test
	public void test2() throws Exception {

		POICellFormatter cellFormatter = new POICellFormatter();
		
		Workbook wb  = WorkbookFactory.create(new File("files/poi.xls"));
		Sheet sheet = wb.getSheetAt(0);
		Sheet sh = wb.createSheet("abc");
		

		int j = 0;
		for(Row row: sheet) {
			Row row2 = sh.createRow(j++);
			int i = 0;
			for (Cell cell : row) {
				Cell cell2 = row2.createCell(i++);
				System.out.print(cellFormatter.formatAsString(cell)+"\t");
				cloneCell(cell2, cell);
				
			}
			System.out.println();
		}
		
		wb.write(new FileOutputStream(new File("files/hoge.xls")));
	}
	
	@Test
	public void test4() throws Exception {

		Workbook workbook  = WorkbookFactory.create(new File("files/aaa.xls"));
		Sheet sheet = workbook.getSheetAt(0);
		Copy copy = new Copy();
		copy.workbook = workbook;
		copy.sheet = sheet;
		// (開始行 startY 123,　開始列 startX abc) (終了行 endY 123,　終了列 endX abc)
		copy.copyCellArea(0, 1, 9, 3, 0, 2);
		workbook.write(new FileOutputStream(new File("files/bbb.xls")));
	}
	
	//列の挿入
	@Test
	public void test5() throws Exception {

		Workbook workbook  = WorkbookFactory.create(new File("files/aaa.xls"));
		Sheet sheet = workbook.getSheetAt(0);
				
		int col = 5;
		sheet.forEach(row -> {
			int last = row.getLastCellNum();
			for (int i = last; i >= col; i--) {
				Cell srcCell = row.getCell(i,   MissingCellPolicy.CREATE_NULL_AS_BLANK);
				Cell dstCell = row.getCell(i+1, MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cloneCell(dstCell, srcCell);
			}
			Cell cell = row.getCell(col, MissingCellPolicy.CREATE_NULL_AS_BLANK);	
			cell.setCellType(Cell.CELL_TYPE_BLANK);
		});
		
		workbook.write(new FileOutputStream(new File("files/bbb.xls")));
	}
	
	//行の挿入
	@Test
	public void test6() throws Exception {

		Workbook workbook  = WorkbookFactory.create(new File("files/aaa.xls"));
		Sheet sheet = workbook.getSheetAt(0);
		
		int rowNum = 0;
		int last = sheet.getLastRowNum();
		if (rowNum <= last) {
			sheet.shiftRows(rowNum, sheet.getLastRowNum(), 1);
		} else {
			sheet.createRow(rowNum);
		}
		
		workbook.write(new FileOutputStream(new File("files/bbb.xls")));
	}
	
	public static class InsertColumn implements Consumer<Row> {
		
		int column;
		
		public InsertColumn(int i) {
			column = i;
		}
		
		@Override
		public void accept(Row row) {
		  Cell cell = row.getCell(column, MissingCellPolicy.CREATE_NULL_AS_BLANK);	
		  cell.setCellValue("NULL");
		}
	}
		
	
	
	private static void cloneCell(Cell cNew, Cell cOld) {
		
		cNew.setCellComment(cOld.getCellComment());
		cNew.setCellStyle(cOld.getCellStyle());

		switch (cOld.getCellType()) {
		case Cell.CELL_TYPE_BOOLEAN: {
			cNew.setCellValue(cOld.getBooleanCellValue());
			break;
		}
		case Cell.CELL_TYPE_NUMERIC: {
			cNew.setCellValue(cOld.getNumericCellValue());
			break;
		}
		case Cell.CELL_TYPE_STRING: {
			cNew.setCellValue(cOld.getStringCellValue());
			break;
		}
		case Cell.CELL_TYPE_ERROR: {
			cNew.setCellValue(cOld.getErrorCellValue());
			break;
		}
		case Cell.CELL_TYPE_FORMULA: {
			cNew.setCellFormula(cOld.getCellFormula());
			break;
		}
		}
	}
	
	
	@Test
	public void test3() throws Exception {
		update(new File("files/a.xls"));
		update(new File("files/a.xlsx"));
	}
	
	public static void update(File file) throws Exception {
		
		InputStream is = new FileInputStream(file);
		Workbook wb = WorkbookFactory.create(is); // メモリ上に展開
		is.close(); // ここで入力ストリームを閉じる

		// 編集したいシート、列、セルを指定
		Sheet s = wb.getSheetAt(0);
		Row r = s.getRow(0);
		Cell c = r.getCell(0);
		
		// この場合B2セルに「あいう」をセット
		c.setCellValue("あいう");
		c.setCellType(Cell.CELL_TYPE_STRING);

		// 編集した内容の書き出し
		OutputStream os = new FileOutputStream(file);
		wb.write(os);
		os.close();
	}	
	
	
	private void setDate(Cell cell, Date date, Workbook wb) {
		CreationHelper createHelper = wb.getCreationHelper();
		CellStyle cellStyle = wb.createCellStyle();
		short style = createHelper.createDataFormat().getFormat("yyyy-mm-dd\"T\"hh:mm\"+09:00\"");
		System.out.println(style);
		cell.setCellValue(date);
		cellStyle.setDataFormat((short)0xE);
		cell.setCellStyle(cellStyle);
	}
	
	
	
	
	
	
	
/*
	0, "General"
1, "0"
2, "0.00"
3, "#,##0"
4, "#,##0.00"
5, "$#,##0_);($#,##0)"
6, "$#,##0_);[Red]($#,##0)"
7, "$#,##0.00);($#,##0.00)"
8, "$#,##0.00_);[Red]($#,##0.00)"
9, "0%"
0xa, "0.00%"
0xb, "0.00E+00"
0xc, "# ?/?"
0xd, "# ??/??"
0xe, "m/d/yy"
0xf, "d-mmm-yy"
0x10, "d-mmm"
0x11, "mmm-yy"
0x12, "h:mm AM/PM"
0x13, "h:mm:ss AM/PM"
0x14, "h:mm"
0x15, "h:mm:ss"
0x16, "m/d/yy h:mm"
// 0x17 - 0x24 reserved for international and undocumented 0x25, "#,##0_);(#,##0)"
0x26, "#,##0_);[Red](#,##0)"
0x27, "#,##0.00_);(#,##0.00)"
0x28, "#,##0.00_);[Red](#,##0.00)"
0x29, "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)"
0x2a, "_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)"
0x2b, "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)"
0x2c, "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)"
0x2d, "mm:ss"
0x2e, "[h]:mm:ss"
0x2f, "mm:ss.0"
0x30, "##0.0E+0"
0x31, "@" - This is text format.
0x31 "text" - Alias for "@"
*/
}
