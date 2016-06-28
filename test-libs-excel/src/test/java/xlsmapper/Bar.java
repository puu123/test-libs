package xlsmapper;

import com.gh.mygreen.xlsmapper.annotation.XlsColumn;

public class Bar {

	@XlsColumn(columnName = "id")
	String name;

	@Override
	public String toString() {
		return "Bar [name=" + name + "]";
	}

}
