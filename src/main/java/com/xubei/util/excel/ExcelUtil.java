package com.xubei.util.excel;

import com.xubei.enums.DatePatternEnum;
import com.xubei.util.lang.DateUtil;
import com.xubei.util.lang.StringUtil;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * 基于poi的excel工具类
 * 相对老版来说 用Workbook、Sheet、Row、Cell等父类 来接收处理就避免了不同后缀写2套的麻烦
 * 还可以使用easyexcel 来实现大数量的情况
 * todo 根据List<xxVO> 导出,如果有时间类型 自动转换为String
 *
 * @author shuheng
 */
public class ExcelUtil {
	
	/**
	 * 读取excel 原理：一行一行读取内容，保存到List<String[]>中
	 * 
	 * @param filePath 	需要读取的文件路径
	 * @return
	 * @throws IOException
	 */
	public static List<String[]> read(String filePath) throws IOException {
		File file = new File(filePath);
		return read(file);
	}
	
	/**
	 * 读取excel 原理：一行一行读取内容，保存到List<String[]>中
	 * 
	 * @param file 	需要读取的流
	 * @return
	 * @throws IOException
	 */
	public static List<String[]> read(File file) throws IOException {
		String fileName = file.getName();
		InputStream inStream = new FileInputStream(file);
		return read(fileName, inStream);
	}
	
	/**
	 * 读取excel 原理：一行一行读取内容，保存到List<String[]>中
	 * 
	 * @param fileName 	文件名,主要是为了后缀名
	 * @param inStream 	需要读取的流
	 * @return
	 * @throws IOException
	 */
	public static List<String[]> read(String fileName, InputStream inStream) throws IOException {
		// 创建一个list 用来存储读取的内容
		List<String[]> list = new ArrayList<String[]>();
		// 前缀prefix，后缀suffix
//		String suffix = filePath.substring(filePath.lastIndexOf(".") + 1);
		String suffix = StringUtil.getSuffix(fileName);
		
		Workbook workbook = createWorkbook(suffix, inStream);
		if(workbook == null){
			inStream.close();
			return list;
		}
		// 默认读取第一个sheet
		// Sheet sheet = workbook.getSheetAt(0);
		int sheetTotal = workbook.getNumberOfSheets();
		for (int numSheet = 0; numSheet < sheetTotal; numSheet++) {
			Sheet sheet = workbook.getSheetAt(numSheet);
			if (sheet == null) {
				continue;
			}
			int rownum = sheet.getLastRowNum();// 总行数
			// Read the Row
			for (int i = 1; i <= rownum; i++) {
				Row row = sheet.getRow(i);
				int colnum = row.getLastCellNum();// 总列数
				String[] str = new String[colnum];
				if (row != null) {
					for (int j = 0; j < colnum; j++) {
						Cell cell = row.getCell(j);
						str[j] = getValue(cell);
					}
					list.add(str);
				}
			}
		}
		inStream.close();
		workbook.close();//关闭
		
		return list;
	}
	
	/**
	 * 根据xls或xlsx 创建相应 Workbook
	 * @param suffix
	 * @param inStream
	 * @return
	 * @throws IOException
	 */
	private static Workbook createWorkbook(String suffix, InputStream inStream) throws IOException{
		Workbook workbook = null;
		if ("xls".equals(suffix)) {
			workbook = new HSSFWorkbook(inStream);
		} else if ("xlsx".equals(suffix)) {
			workbook = new XSSFWorkbook(inStream);
		}
		
		return workbook;
	}
	
	/**
	 * 根据xls或xlsx 创建相应 Workbook
	 * @param suffix
	 * @return
	 * @throws IOException
	 */
	private static Workbook createWorkbook(String suffix){
		Workbook workbook = null;
		if ("xls".equals(suffix)) {
			workbook = new HSSFWorkbook();
		} else if ("xlsx".equals(suffix)) {
			workbook = new XSSFWorkbook();
		}
		
		return workbook;
	}
	
	/**
	 * 获取单元格内容
	 * 
	 * @param cell
	 * @return
	 */
	/**
	 * 获取单元格内容
	 *
	 * @param cell
	 * @return
	 */
	private static String getValue(Cell cell) {
		//兼容null 和空格的情况
		if(cell == null){
			return "";
		}

		String val = "";
		switch (cell.getCellType()) {
			case Cell.CELL_TYPE_NUMERIC:
                val = getNumberValue(cell);
				break;
			case Cell.CELL_TYPE_STRING: //文本类型
				val = cell.getStringCellValue();
				break;
			case Cell.CELL_TYPE_BOOLEAN: //布尔型
				val = String.valueOf(cell.getBooleanCellValue());
				break;
			case Cell.CELL_TYPE_BLANK: //空白
				val = "";
				break;
			case Cell.CELL_TYPE_ERROR: //错误
				val = "错误";
				break;
			case Cell.CELL_TYPE_FORMULA: //公式
				try {
					val = String.valueOf(cell.getStringCellValue());
				} catch (IllegalStateException e) {
					val = String.valueOf(cell.getNumericCellValue());
				}
				break;
			default:
                RichTextString richStringCellValue = cell.getRichStringCellValue();
                val = richStringCellValue == null ? "" : richStringCellValue.toString();
		}
		//去除所有空格
		if(val != null){
			val = val.replace(" ", "");
		}
		return val;
//        String val = "";
//		int cellType = cell.getCellType();
//		if (cellType == Cell.CELL_TYPE_BLANK) {
//            val = "";
//		}else if (cellType == Cell.CELL_TYPE_BOOLEAN) {
//            val = String.valueOf(cell.getBooleanCellValue());
//		} else if (cellType == Cell.CELL_TYPE_STRING) {
//            val = cell.getStringCellValue();
//        } else if (cellType == Cell.CELL_TYPE_NUMERIC) {
//            val = getNumberValue(cell);
//        } else {
//			return cell.getStringCellValue();
//		}
//        return val;
	}

	//获得单元格是数字类型的值
    private static String getNumberValue(Cell cell) {
        String val = "";
        //HSSFDateUtil底层是org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted()
        if(HSSFDateUtil.isCellDateFormatted(cell)){
            val = DateUtil.dateToStr(cell.getDateCellValue(), DatePatternEnum.YMD.getCode());
        }else{
            double numericCellValue = cell.getNumericCellValue();
            if(numericCellValue % 1 == 0){//整数
                DecimalFormat df = new DecimalFormat("0");
                val = df.format(numericCellValue); //数字型
            }else {//小数
                val = String.valueOf(numericCellValue);
            }
        }
        return val;
    }
	
	/**
	 * 导出excel (推荐)
	 * 
	 * @param excelInfo
	 * @param outStream
	 * @throws IOException
	 */
	public static void write(ExcelInfo excelInfo, OutputStream outStream) throws IOException {
//		String fileName = excelInfo.getFileName();
//		// 前缀prefix，后缀suffix
////		String suffix = fileName.substring(fileName.lastIndexOf(".") + 1);
//		String suffix = StringUtil.getSuffix(fileName);
//		// 创建Workbook对象(excel的文档对象)
//		Workbook workbook = createWorkbook(suffix);
//		// 建立新的sheet对象（excel的表单）
//		Sheet sheet = workbook.createSheet(excelInfo.getSheetName());
////		//设置缺省列宽8.5,行高为设置的20
////		sheet.setDefaultRowHeightInPoints(20);
//
//		String[] titles = excelInfo.getTitles();
//		String[] fields = excelInfo.getFields();
//		setTitleRow(sheet, titles);
//		
//		//具体内容
//		List<Map<String, Object>> list = excelInfo.getMapList();
//		if(list != null){
//			int size = list.size();
//			for (int i = 0; i < size; i++) {
//				setContentRows(sheet, i+1, list.get(i), fields, titles);
//			}
//		}
//		workbook.write(outStream);
//		outStream.close();
		
		write(excelInfo.getTitles(), excelInfo.getFields(), excelInfo.getMapList(), excelInfo.getSheetName(),
				excelInfo.getFileName(), outStream);
	}
	
	/**
	 * 数据是mapList 的情况, 每列宽度固定width 长度
	 * 
	 * @param titles
	 * @param fields
	 * @param mapList
	 * @param sheetName
	 * @param fileName
	 * @param outStream
	 * @throws IOException
	 */
	public static <T> void write(String[] titles, String[] fields, List<Map<String, Object>> mapList, String sheetName, String fileName, OutputStream outStream) throws IOException {
		String suffix = StringUtil.getSuffix(fileName);
		// 创建Workbook对象(excel的文档对象)
		Workbook workbook = createWorkbook(suffix);
		// 建立新的sheet对象（excel的表单）
		Sheet sheet = workbook.createSheet(sheetName);
//		//设置缺省列宽8.5,行高为设置的20
//		sheet.setDefaultRowHeightInPoints(20);
		setTitleRow(sheet, titles);

		//具体内容
		if(mapList != null){
			int size = mapList.size();
			for (int i = 0; i < size; i++) {
				setContentRows(sheet, i+1, mapList.get(i), fields, titles);
			}
		}
		workbook.write(outStream);
		outStream.close();
	}
	
	/**
	 * 数据是arrList 的情况, 每列宽度固定width 长度
	 * 
	 * @param titles
	 * @param arrList
	 * @param sheetName
	 * @param fileName
	 * @param width  只用填20之类的数, 基数已添加
	 * @param outStream
	 * @throws IOException
	 */
	public static <T> void write(String[] titles, List<T[]> arrList, String sheetName, String fileName, int width, OutputStream outStream) throws IOException {
		String suffix = StringUtil.getSuffix(fileName);
		// 创建Workbook对象(excel的文档对象)
		Workbook workbook = createWorkbook(suffix);
		// 建立新的sheet对象（excel的表单）
		Sheet sheet = workbook.createSheet(sheetName);
//		//设置缺省列宽8.5,行高为设置的20
//		sheet.setDefaultRowHeightInPoints(20);
		setTitleRow(sheet, titles);

		//（汉字是512，数字是256）, 统一乘以 512
		int trueWidth = width * 512;
		sheet.setDefaultColumnWidth(trueWidth);
		//具体内容
		if(arrList != null){
			int size = arrList.size();
			for (int i = 0; i < size; i++) {
				setContentRowsFixexWidth(sheet, i+1, arrList.get(i));
			}
		}
		workbook.write(outStream);
		outStream.close();
	}
	
	/**
	 * 数据是arrList 的情况, 每列宽度固定width 长度
	 * 
	 * @param titles
	 * @param arrList
	 * @param sheetName
	 * @param fileName
	 * @param outStream
	 * @throws IOException
	 */
	public static <T> void write(String[] titles, List<T[]> arrList, String sheetName, String fileName, OutputStream outStream) throws IOException {
		String suffix = StringUtil.getSuffix(fileName);
		// 创建Workbook对象(excel的文档对象)
		Workbook workbook = createWorkbook(suffix);
		// 建立新的sheet对象（excel的表单）
		Sheet sheet = workbook.createSheet(sheetName);
//		//设置缺省列宽8.5,行高为设置的20
//		sheet.setDefaultRowHeightInPoints(20);
		setTitleRow(sheet, titles);

		//具体内容
		if(arrList != null){
			int size = arrList.size();
			for (int i = 0; i < size; i++) {
				setContentRowsCalcWidth(sheet, i+1, arrList.get(i), titles);
			}
		}
		workbook.write(outStream);
		outStream.close();
	}
	
	/**
	 * 导出excel,并分页
	 * @param excelInfo
	 * @param outStream
	 * @throws IOException
	 */
	public static void write2Page(ExcelInfo excelInfo, OutputStream outStream) throws IOException {
		String fileName = excelInfo.getFileName();
		// 前缀prefix，后缀suffix
//		String suffix = fileName.substring(fileName.lastIndexOf(".") + 1);
		String suffix = StringUtil.getSuffix(fileName);
		// 创建Workbook对象(excel的文档对象)
		Workbook workbook = createWorkbook(suffix);
		
		//具体内容
		List<Map<String, Object>> list = excelInfo.getMapList();
		int rows = list.size();
		int pageSize = excelInfo.getPageSize();
		int sheetNum = 0; //指定sheet的页数

		if (rows % pageSize == 0) {
			sheetNum = rows / pageSize;
		} else {
			sheetNum = rows / pageSize + 1;
		}
		
		for (int i = 1; i <= sheetNum; i++) {
			// 建立新的sheet对象（excel的表单）
			Sheet sheet = workbook.createSheet(excelInfo.getSheetName() + i);
//			// 设置缺省列宽8.5,行高为设置的20
//			sheet.setDefaultRowHeightInPoints(20);
			
			String[] titles = excelInfo.getTitles();
			String[] fields = excelInfo.getFields();
			setTitleRow(sheet, titles);
			
			//分页处理内容
			int end = Math.min(pageSize, rows);
			for (int s = 0; s < end; s++) {
				int pageCount = (i -1) * pageSize;
				int start = pageCount + s;
				if (start >= rows){
					//如果数据超出总的记录数的时候，就退出循环
					break;
				}
				setContentRows(sheet, s+1, list.get(start), fields, titles);
			}
		}
		workbook.write(outStream);
		outStream.close();
	}
	
	/**
	 * 浏览器上导出excel,并解决文件名 中文乱码
	 * 
	 * @param excelInfo
	 * @param response
	 * @param isNeedPage 是否需要分sheet页
	 * @throws IOException
	 */
	public static void export2Http(ExcelInfo excelInfo, HttpServletResponse response, boolean isNeedPage) throws IOException {
		String fileName = excelInfo.getFileName();

		response.reset();
//		response.setHeader("Content-disposition", "attachment; filename=" + fileName);
		response.setHeader("Content-disposition", "attachment; filename=" + new String(fileName.getBytes("UTF-8"), "ISO8859-1"));
		response.setContentType("application/msexcel");
		OutputStream output = response.getOutputStream();
		
		if(isNeedPage){
			write2Page(excelInfo, output);
		}else{
			write(excelInfo, output);
		}
	}

	/**
	 * 填充标题行
	 * @param sheet
	 * @param titles
	 */
	private static void setTitleRow(Sheet sheet, String[] titles) {
		// 在sheet里创建第一行，参数为行索引(excel的行)，可以是0～65535之间的任何一个
		Row header = sheet.createRow(0);
		header.setHeightInPoints(25);
		// 创建单元格并设置单元格内容
		for (int i = 0, max = titles.length; i < max; i++) {
			header.createCell(i).setCellValue(titles[i]);
		}
	}
	
	/**
	 * 填充内容行
	 * @param sheet
	 * @param rowNum 行号
	 * @param map	对象
	 * @param fields
	 * @param titles
	 */
	private static void setContentRows(Sheet sheet, int rowNum, Map<String, Object> map, String[] fields, String[] titles){
		//sheet.autoSizeColumn(i,true);//中文还是不能实现宽度自适应
		Row row = sheet.createRow(rowNum);
		row.setHeightInPoints(20);
		for (int j = 0, max = fields.length; j < max; j++) {
			String value = String.valueOf(map.get(fields[j]));
			String titleValue = titles[j];
			int valueLength = value.getBytes().length;
			int titleValueLength = titleValue.getBytes().length;
//			int cellLength = (valueLength >= titleValueLength) ? valueLength : titleValueLength;
			int cellLength = Math.max(valueLength, titleValueLength);
			
			//由于utf-8一个中文字符有3个长度
			sheet.setColumnWidth(j,(cellLength + 3) * 256);//手动设置列宽
			row.createCell(j).setCellValue(value);
//			Cell cell = row.createCell(j);
//			cell.setCellType(Cell.CELL_TYPE_STRING);
//			cell.setCellValue(value);
		}
	}

	/**
	 * 每列设置一样的宽度（适合title的文字过多的情况）
	 * 
	 * @param sheet
	 * @param rowNum
	 * @param objArr
	 */
	private static void setContentRowsFixexWidth(Sheet sheet, int rowNum, Object[] objArr){
		//sheet.autoSizeColumn(i,true);//中文还是不能实现宽度自适应
		//只有填20之类的数, 基数已添加（汉字是512，数字是256）, 统一乘以 512
		//由于这个在循环内被提出
//		int trueWidth = 25 * 512;
//		sheet.setDefaultColumnWidth(trueWidth);
		Row row = sheet.createRow(rowNum);
		row.setHeightInPoints(20);
		for (int j = 0, max = objArr.length; j < max; j++) {
			String value = String.valueOf(objArr[j]);
			row.createCell(j).setCellValue(value);
//			Cell cell = row.createCell(j);
//			cell.setCellType(Cell.CELL_TYPE_STRING);
//			cell.setCellValue(value);
		}
	}
	
	/**
	 * 设置单元格的值, 且每列的宽度自动计算
	 * @param sheet
	 * @param rowNum
	 * @param objArr
	 * @param titles
	 */
	private static void setContentRowsCalcWidth(Sheet sheet, int rowNum, Object[] objArr, String[] titles){
		//sheet.autoSizeColumn(i,true);//中文还是不能实现宽度自适应
		Row row = sheet.createRow(rowNum);
		row.setHeightInPoints(20);
		for (int j = 0, max = titles.length; j < max; j++) {
			String value = String.valueOf(objArr[j]);

			String titleValue = titles[j];
			int valueLength = value.getBytes().length;
			int titleValueLength = titleValue.getBytes().length;
//			int cellLength = (valueLength >= titleValueLength) ? valueLength : titleValueLength;
			int cellLength = Math.max(valueLength, titleValueLength);
			//由于utf-8一个中文字符有1个长度
			sheet.setColumnWidth(j,(cellLength + 1) * 256);//手动设置列宽
			row.createCell(j).setCellValue(value);
//			Cell cell = row.createCell(j);
//			cell.setCellType(Cell.CELL_TYPE_STRING);
//			cell.setCellValue(value);
		}
	}

}
