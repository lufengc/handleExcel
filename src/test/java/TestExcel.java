import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.read.metadata.ReadSheet;

/**
 *
 * @author fengcheng
 * @version 2020/12/10
 */
public class TestExcel {

	public static String inPath = "exec/src.xlsx";
	public static String outPath = "exec/src.xlsx";

	public static void main(String[] args) {
		ExcelReader excelReader = EasyExcel.read(inPath).build();
		excelReader.read(readSheet(0));
		excelReader.finish();
	}

	private static ReadSheet readSheet(int i) {
		return EasyExcel.readSheet(i).head(ExcelDto.class).registerReadListener(new TestExcelListener()).build();
	}

}
