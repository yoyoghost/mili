import java.io.IOException;
import java.io.OutputStream;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletResponse;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Excel导出工具类，利用poi实现
 * @author jojo
 *
 */
public class ExportExcelUtil {

	/**
	 * 工作薄对象
	 */
	private SXSSFWorkbook wb;
	
	/**
	 * 工作表对象
	 */
	private Sheet sheet;
	
	/**
	 * 当前操作的是第几行
	 */
	private int rownum=0;
	
	/**
	 * 得到要导出的excel的列名称
	 * 
	 * @param cls
	 * @return
	 */
	public String[] getExcelTitle(Class<?> cls) {
		// 利用反射机制得到属性字段
		Field[] excelField = cls.getDeclaredFields();
		String[] excelFieldArray = new String[excelField.length];
		for (int k = 0; k < excelField.length; k++) {
			Field field = excelField[k];
			// 得到属性上面的注解
			Annotation annotation = field.getAnnotation(FieldMeta.class);
			FieldMeta fieldMeta = (FieldMeta) annotation;
			// 得到注解字段对应的值
			excelFieldArray[k] = fieldMeta.aliasName();
		}
		return excelFieldArray;
	}

	/**
	 * 得到excel中要填充的值
	 * @param cls
	 * @param list
	 * @return
	 */
	public Map<Integer, List<Object>> getExcelContent(Class<?> cls,
			List<?> list) {
		// 存储要填充的值（所有行和列）用map是为了保证排序
		Map<Integer, List<Object>> result = new HashMap<Integer, List<Object>>();
		// 存储每行的值
		List<Object> resultList = null;
		// 得到excel中填充字段的值
		for (int i = 0; i < list.size(); i++) {
			Object bean = list.get(i);
			cls = bean.getClass();
			Field[] valueField = cls.getDeclaredFields();
			resultList = new ArrayList<Object>();
			for (int k = 0; k < valueField.length; k++) {
				Field tempField = valueField[k];
				String feildName = tempField.getName();
				// 得到get方法的名字
				String getterMethodName = "get" + StringUtils.capitalize(feildName);
				try {
					Method getMethod = cls.getMethod(getterMethodName, null);
					resultList.add(getMethod.invoke(bean, null));
				} catch (NoSuchMethodException e) {
					e.printStackTrace();
				} catch (SecurityException e) {
					e.printStackTrace();
				} catch (IllegalAccessException e) {
					e.printStackTrace();
				} catch (IllegalArgumentException e) {
					e.printStackTrace();
				} catch (InvocationTargetException e) {
					e.printStackTrace();
				}
			}
			result.put(i, resultList);
			resultList = null;
		}

		return result;
	}
	
	
	/**
	 * 初始化函数
	 * @param title 表格标题，传“空值”，表示无标题（可为空）
	 * @param sheetName (可为空)
	 * @param excelTitle 表头列表
	 */
	public void initialize(String title,String sheetName,String[] excelTitle) {
		this.wb = new SXSSFWorkbook();
		//sheet名字
		if (StringUtils.isNotBlank(sheetName)){
			this.sheet = wb.createSheet(sheetName);
		}else{
			this.sheet = wb.createSheet();
		}
		//创建标题
		if (StringUtils.isNotBlank(title)){
			Row titleRow = sheet.createRow(rownum++);
			titleRow.setHeightInPoints(30);
			Cell titleCell = titleRow.createCell(0);
			titleCell.setCellValue(title);
			sheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(),titleRow.getRowNum(), titleRow.getRowNum(), excelTitle.length-1));
		}
		//创建表头
		if (excelTitle == null){
			throw new RuntimeException("headerList not null!");
		}
		Row headerRow = sheet.createRow(rownum++);
		headerRow.setHeightInPoints(16);
		for (int i = 0; i < excelTitle.length; i++) {
			Cell cell = headerRow.createCell(i);
			cell.setCellValue(excelTitle[i]);
			sheet.autoSizeColumn(i, true);
		}
		for (int j = 0; j < excelTitle.length; j++) {  
			int colWidth = sheet.getColumnWidth(j)*2;
	        sheet.setColumnWidth(j, colWidth < 3000 ? 3000 : colWidth);  
		}
		log.debug("Initialize success.");
	}
	
	/**
	 * 写入数据
	 * @param contentMap
	 */
	public void setContent(Map<?,List<Object>> contentMap){
		for (Map.Entry<?, List<Object>> entry : contentMap.entrySet()) {
			Row row = sheet.createRow(rownum++);
			List<Object> value = entry.getValue();
			for (int i = 0; i < value.size(); i++) {
				Cell cell = row.createCell(i);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				if(null==value.get(i)){
					cell.setCellValue("");
				}else{
					cell.setCellValue(value.get(i)+"");
				}
			}
		}
	}
	
	/**
	 * 输出excel
	 * @param response
	 * @param fileName
	 */
	public void write(HttpServletResponse response, String fileName){
        try {
        	response.reset();
        	response.setCharacterEncoding("UTF-8");
            response.setContentType("application/octet-stream; charset=utf-8");
            response.setHeader("Content-Disposition", "attachment; filename="+Encodes.urlEncode(fileName));
			OutputStream os = response.getOutputStream();
			wb.write(os);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * 关闭
	 */
	public void dispose(){
		wb.dispose();
	}
}
