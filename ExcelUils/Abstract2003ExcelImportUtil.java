package test.excelutils;

import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;

public abstract class Abstract2003ExcelImportUtil extends Abstract2003ExcelUtil {

	@Override
	@Deprecated
	protected void createTitle(HSSFSheet sheet) throws Exception {
		throw new RuntimeException("该方法未被实现，请调用Abstract2003ExcelExportUtil类的createTitle方法!");
	}

	@Override
	@Deprecated
	protected void createExcelData(HSSFRow row, List<Object> data, int j)
			throws Exception {
		throw new RuntimeException("该方法未被实现，请调用Abstract2003ExcelExportUtil类的createExcelData方法!");
	}


}
