package test.excelutils;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import junit.framework.TestCase;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class TestExcelUtilClasses extends TestCase {

	
	/**
	 * 导出2003excel例子
	 * @throws Exception
	 */
	public void testExportExcel() throws Exception{
		Abstract2003ExcelExportUtil util = new Abstract2003ExcelExportUtil() {
			
			@Override
			protected void createTitle(HSSFSheet sheet) throws Exception {
				HSSFRow row = sheet.createRow(0);
				row.createCell(0).setCellValue("第一列");
				row.createCell(1).setCellValue("第一列");
				row.createCell(2).setCellValue("第二列");
			}
			
			@Override
			protected void createExcelData(HSSFRow row, List<Object> data, int j)
					throws Exception {
				@SuppressWarnings("unchecked")
				Map<String,String> map = (Map<String,String>)data.get(j);
				row.createCell(0).setCellValue(map.get("第一列"));
				row.createCell(1).setCellValue(map.get("第二列"));
				row.createCell(2).setCellValue(map.get("第三列"));
			}
		};
		
		List<Object> data = new ArrayList<Object>();
		for(int i=0; i<10;i++){
			Map<String,String> temp  = new HashMap<String, String>();
			temp.put("第一列", 1*(i+1)+"");
			temp.put("第二列", 2*(i+1)+"");
			temp.put("第三列", 3*(i+1)+"");
			data.add(temp);
		}
		HSSFWorkbook workbook = util.createHSSFWorkbook(data);
		ExcelUtils.saveExportFile("d:/test/", "test.xls", workbook);
	}
	
	/**
	 * 导入2003excel例子
	 * @throws Exception
	 */
	public void testImportExcel() throws Exception{
		Abstract2003ExcelImportUtil util = new Abstract2003ExcelImportUtil() {
			
			@Override
			protected Object parseDataToObject(HSSFRow row) throws Exception {
				Map<String,Object> temp = new HashMap<String, Object>();
				String value0 = ExcelUtils.getValue(row.getCell(0));
				String value1 = ExcelUtils.getValue(row.getCell(1));
				String value2 = ExcelUtils.getValue(row.getCell(2));
				temp.put("第一列", value0);
				temp.put("第二列", value1);
				temp.put("第三列", value2);
				return temp;
			}
		};
		
		List<Object> data = util.importData(new File("d:/test/test.xls"), 1);
		System.out.println(data.toString());
	}
}
