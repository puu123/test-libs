package xlsbeans;

import java.util.Date;
import java.util.List;

import net.java.amateras.xlsbeans.annotation.HorizontalRecords;
import net.java.amateras.xlsbeans.annotation.LabelledCell;
import net.java.amateras.xlsbeans.annotation.LabelledCellType;
import net.java.amateras.xlsbeans.annotation.Sheet;

@Sheet(number=0)
public class Foo {
	
	@LabelledCell(label="タイトル",type=LabelledCellType.Right)
	public String title;
	
	@LabelledCell(label="更新日",type=LabelledCellType.Right)
	public String updateAt;
	
	@HorizontalRecords(headerRow=4,headerColumn=0)
	public List<Bar> list;

	@Override
	public String toString() {
		final int maxLen = 10;
		return "Foo [title=" + title + ", updateAt=" + updateAt + ", list="
				+ (list != null ? list.subList(0, Math.min(list.size(), maxLen)) : null) + "]";
	}


	
}
