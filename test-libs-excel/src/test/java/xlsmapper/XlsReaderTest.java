package xlsmapper;

import java.io.FileInputStream;

import org.junit.Test;

import com.gh.mygreen.xlsmapper.XlsMapper;

public class XlsReaderTest {

	@Test
	public void test() throws Exception {

		XlsMapper xlsMapper = new XlsMapper();
		// シートの読み込み
		//xlsmapper xlsMapper = new XlsMapper();
		Foo sheet = xlsMapper.load(
				new FileInputStream("files/xlsmapper/foo.xls"), // 読み込むExcelファイル。
				Foo.class // シートマッピング用のPOJOクラス。
		);
		
		System.out.println(sheet);
	}

}
