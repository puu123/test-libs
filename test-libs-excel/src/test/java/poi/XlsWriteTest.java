package poi;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.junit.Before;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

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