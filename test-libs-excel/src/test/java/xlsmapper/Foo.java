package xlsmapper;

import java.util.List;

import com.gh.mygreen.xlsmapper.annotation.XlsHorizontalRecords;
import com.gh.mygreen.xlsmapper.annotation.XlsSheet;

@XlsSheet(number=0)
public class Foo {

	@XlsHorizontalRecords(headerRow=4, headerColumn=0)
	List<Bar> list;

	@Override
	public String toString() {
		return "Foo [list=" + list + "]";
	}
	
	
}
