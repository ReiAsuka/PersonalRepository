package test.excelutils;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.formula.functions.T;

public abstract class Abstract2003ExcelExportUtil extends Abstract2003ExcelUtil {

	@Deprecated
	@Override
	protected T parseDataToObject(HSSFRow row) throws Exception {
		throw new RuntimeException("该方法未被实现，请调用Abstract2003ExcelImportUtil类的parseDataToObject方法!");
	}

}
