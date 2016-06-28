package xlsbeans;

import com.gh.mygreen.xlsmapper.annotation.XlsColumn;

import net.java.amateras.xlsbeans.annotation.Column;

public class Bar {

	@Column(columnName = "id")
	public String name;

	@Override
	public String toString() {
		return "Bar [name=" + name + "]";
	}

}
