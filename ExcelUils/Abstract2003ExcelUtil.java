package test.excelutils;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * 这个类是抽象的excel表格2003工具类，基于poi框架，
 * 有创建excel（HSSFWorkbook）对象的方法createHSSFWorkbook，
 * 有导入excel数据的方法importData
 * @author cy
 *
 */
public abstract class Abstract2003ExcelUtil {
	/**
	 * 该抽象方法用于创建excel第一排（或者前面几排）的标题
	 * @param sheet 当前sheet表，可以使用sheet.createRow(0)方法创建出多个row（HSSFRow）
	 */
	protected abstract void createTitle(HSSFSheet sheet) throws Exception;
	/**
	 * 该抽象方法用于创建标题下面的数据
	 * @param row 当前行
	 * @param data 填充excel的数据集合
	 * @param j 当前行应该填充数据集合（data）中第j个索引对象的数据
	 */
	protected abstract void createExcelData(HSSFRow row,List<Object> data,int j) throws Exception;
	/**
	 * 创建HSSFWorkbook对象，根据数据集合的大小创建多个sheet，然后调用抽象方法createTitle设置title，
	 * 创建装数据的row，然后调用抽象方法createExcelData填充数据，并返回HSSFWorkbook对象
	 * @param data
	 * @return
	 * @throws Exception
	 */
	public HSSFWorkbook createHSSFWorkbook(List<Object> data) throws Exception{
		HSSFWorkbook workbook = new HSSFWorkbook();
		if (data != null && data.size()>0){
			for(int i = 0; i<=(data.size()/65535);i++){
				HSSFSheet sheet = workbook.createSheet("Sheet"+(i+1));
	        	this.createTitle(sheet);
	        	int row=1;
	        	int nextSheetLimit = (i+1)*65535;
	        	int currentLimit = nextSheetLimit>data.size()?data.size():nextSheetLimit;
	        	for(int j = i*65535; j<currentLimit;j++) 	{
	        		HSSFRow rowData = sheet.createRow(row);
	        		this.createExcelData(rowData, data,j);
	        		row++;
	        	}
			}
        }
		return workbook;
	}
	
	
	
	/**
	 * 读取excel里的数据,调用抽象方法parseDataToObject(HSSFRow row)把每行数据解析成具体的对象，
	 * 需要实现该抽象方法
	 * @param file 需要导入数据的.xls文件
	 * @param startRowNum 从startRowNum行开始读取数据，0表示是第一行
	 * @return 返回具体的对象的集合
	 * @throws Exception
	 */
	public List<Object> importData(File file,int startRowNum) throws Exception{
		String fileName = file.getName();
		//获得原始文件后缀名
		String fileAttr = fileName.substring(fileName.lastIndexOf(".")+1);
		if(!fileAttr.equals("xls"))
			return null;
		List<Object> result = new ArrayList<Object>();
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook(new FileInputStream(file));

		
		//遍历每个sheet
		for (int i = 0; i < hssfWorkbook.getNumberOfSheets(); i++) {
			HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(i);
			if(hssfSheet!=null){
			//遍历每行记录，由于getLastRowNum得到的行数从0开始，因此得到的num比实际行数少1,所以变成了<=
				for (int j = startRowNum; j <= hssfSheet.getLastRowNum(); j++) {
					HSSFRow row = hssfSheet.getRow(j);
					Object temp = parseDataToObject(row);
					result.add(temp);
				}
			}
		}
		return result;
	}
	
	/**
	 * 解析每行对象，封装成具体的对象
	 * @param row 每行数据
	 * @return 封装好的具体的对象
	 * @throws Exception 
	 */
	protected abstract Object parseDataToObject(HSSFRow row) throws Exception;
}
