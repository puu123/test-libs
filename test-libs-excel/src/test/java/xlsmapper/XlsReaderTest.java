package xlsmapper;

import java.io.File;
import java.io.FileInputStream;

import org.junit.Test;

import com.gh.mygreen.xlsmapper.XlsMapper;
import com.gh.mygreen.xlsmapper.annotation.XlsSheet;
import com.gh.mygreen.xlsmapper.xml.AnnotationReader;
import com.gh.mygreen.xlsmapper.xml.XmlIO;
import com.gh.mygreen.xlsmapper.xml.bind.XmlInfo;


public class XlsReaderTest {

	@Test
	public void test() throws Exception {

		XlsMapper xlsMapper = new XlsMapper();
		// シートの読み込み
		Foo sheet = xlsMapper.load(
				new FileInputStream("files/xlsmapper/foo.xls"), // 読み込むExcelファイル。
				Foo.class // シートマッピング用のPOJOクラス。
		);
		
		System.out.println(sheet);
	}
	
	@Test
	public void test2() throws Exception {
		
//		// XMLファイルの読み込み
//		XmlInfo xmlInfo = XmlIO.load(new File("files/xlsmapper/abc.xml"), "UTF-8");
//
//		// AnnotationReaderのインスタンスを作成
//		AnnotationReader reader = new AnnotationReader(xmlInfo);
//
//		// SheetObjectクラスに付与されたSheetアノテーションを取得
//		XlsSheet sheet = reader.getAnnotation(Hoge.class, XlsSheet.class);
//		System.out.println(sheet.number());

		XlsMapper xlsMapper = new XlsMapper();
		// シートの読み込み
		Hoge sheet = xlsMapper.load(
				new FileInputStream("files/xlsmapper/foo.xls"), // 読み込むExcelファイル。
				Hoge.class, // シートマッピング用のPOJOクラス。
				new FileInputStream("files/xlsmapper/abc.xml")
		);
		
		System.out.println(sheet);
	}
}
