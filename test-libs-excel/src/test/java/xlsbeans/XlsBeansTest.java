package xlsbeans;

import java.io.FileInputStream;
import java.util.Date;

import org.junit.Test;


import net.java.amateras.xlsbeans.XLSBeans;
import net.java.amateras.xlsbeans.XLSBeansConfig;

public class XlsBeansTest {

	@Test
	public void test() throws Exception {

		XLSBeans xlsbeans = new XLSBeans();
		
		XLSBeansConfig config = new XLSBeansConfig();
		config.setNormalizeLabelText(true);
		config.setTrimText(true);
		//config.addTypeConverter(Date.class, new DateTypeConverter());
		xlsbeans.setConfig(config);
		
		// シートの読み込み
		Foo sheet = xlsbeans.load(
				new FileInputStream("files/xlsmapper/foo.xls"), // 読み込むExcelファイル。
				Foo.class // シートマッピング用のPOJOクラス。
		);
		
		System.out.println(sheet);
	}
	
	@Test
	public void test2() throws Exception {

		XLSBeans xlsbeans = new XLSBeans();
		
		XLSBeansConfig config = new XLSBeansConfig();
		config.setNormalizeLabelText(true);
		config.setTrimText(true);
		//config.addTypeConverter(Date.class, new DateTypeConverter());
		xlsbeans.setConfig(config);
		
		// シートの読み込み
		Hoge sheet = xlsbeans.load(
				new FileInputStream("files/xlsmapper/foo.xls"), // 読み込むExcelファイル。
				new FileInputStream("files/xlsmapper/xyz.xml"),
				Hoge.class // シートマッピング用のPOJOクラス。
				
		);
		
		System.out.println(sheet);
	}
}
