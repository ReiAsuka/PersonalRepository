package test.excelutils;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.math.BigDecimal;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

public class ExcelUtils {

	/**
	 * 获取单元格内容
	 * @param hssfCell 单元格
	 * @return 如果单元格为空，则返回""
	 */
	public static String getValue(HSSFCell hssfCell) { 
		if(hssfCell == null)
			return "";
        if (hssfCell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {  
            // 返回布尔类型的值  
            return String.valueOf(hssfCell.getBooleanCellValue());  
        } else if (hssfCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {  
        	BigDecimal bd = new BigDecimal(hssfCell.getNumericCellValue()); 
            // 返回数值类型的值  
            return bd.toPlainString();  
        } else {  
            // 返回字符串类型的值  
            return String.valueOf(hssfCell.getStringCellValue());  
        }  
	}
	
	
	
	/**
	 * 将2003Excel的HSSFWorkbook对象写到磁盘上
	 * @param path 目的路径,如果路径不存在则会创建该路径
	 * @param fileName 文件名
	 * @param hssfworkbook excel对象
	 * @return
	 * @throws Exception 
	 */
	public static String saveExportFile(String path,String fileName,HSSFWorkbook hssfworkbook)
			throws Exception{
		//如果路径不存在则会创建该路径
		File dir = new File(path);
		if(!dir.exists())
			dir.mkdirs();
		
		
		//写入文件
        OutputStream  file = new FileOutputStream (path + File.separator + fileName);
        hssfworkbook.write(file);
        file.close();
        return path + File.separator + fileName;
	}
}
