package weixinkeji.vip.expand.poi;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class JWEOfficeW {
	private Workbook wb=null;
	private OutputStream stream=null;
	public JWEOfficeW(OutputStream stream,JWEOfficeEnum xls_xlsx) throws IOException{
		this.wb = xls_xlsx==JWEOfficeEnum.xls?new HSSFWorkbook():new XSSFWorkbook();
		this.stream=stream;
	}
	public JWEOfficeW(String filePath,JWEOfficeEnum xls_xlsx) throws IOException{
		this.wb = xls_xlsx==JWEOfficeEnum.xls?new HSSFWorkbook():new XSSFWorkbook();
		stream = new FileOutputStream(filePath);
	}
	/**
	 * 取得 Workbook 接口的实例。
	 * @return Workbook
	 */
	public Workbook getWorkbook() {
		return this.wb;
	}
	
	public void write() throws IOException {
		if(null!=this.wb) {
			this.wb.write(this.stream);
		}
	}
	
	/**
	 * 写出文档，与自动关闭输出流
	 * @throws IOException io流异常
	 */
	public void writeAndAutoCloseIO() throws IOException {
		try {
			if(null!=this.wb) {
				this.wb.write(this.stream);
			}
		}catch (Exception e) {
			throw e;
		}finally {
			if(null!=this.stream) {
				this.stream.close();
			}
		}
	}
	
	
	/**
	 * 将集合中的数据，通过输出流，写到excel文档中
	 * @param tablename excel工作表名称
	 * @param list 数据集合
	 * @return boolean  真假
	 * @throws Exception  异常
	 */
	public boolean addToExcel(String tablename, List<?> list) throws Exception {
		if (null == list || list.isEmpty()) {
			return false;
		}
		if(list.isEmpty()) {
			return false;
		}
		Class<?> c = list.get(0).getClass();
		JWEOfficeVO[] vos=null;
		try {
			vos = JWEOfficeVO.getJWEOfficeVO(c);
		} catch (Exception e1) {
			e1.printStackTrace();
			return false;
		}
		// 创建一张表格
		Sheet sheet = wb.createSheet(null == tablename || tablename.isEmpty() ? "sheet1" : tablename);
		// 创建第1行
		Row row = sheet.createRow(0);
		// 设置第1行的数据
		W_PoiOffice.setTitle(row, vos);
		
		// 第二行开始，填充数据
		int j = 1;
		for (int i = 0; i < list.size(); i++) {
			row = sheet.createRow(j++);
			W_PoiOffice.setData(wb, row, vos, list.get(i));
		}
		return true;
	}
	/**
	 * 将集合中的数据，通过输出流，写到excel文档中
	 * @param sheetIndex 第几个excel工作表 （程序组装： "sheet"+sheetIndex ）
	 * @param list 数据集合
	 * @return boolean  真假
	 * @throws Exception  异常
	 */
	public boolean addToExcel(int sheetIndex, List<?> list) throws Exception {
		if (null == list || list.isEmpty()) {
			return false;
		}
		if(list.isEmpty()) {
			return false;
		}
		Class<?> c = list.get(0).getClass();
		JWEOfficeVO[] vos=null;
		try {
			vos = JWEOfficeVO.getJWEOfficeVO(c);
		} catch (Exception e1) {
			e1.printStackTrace();
			return false;
		}
		// 创建一张表格
		Sheet sheet = wb.createSheet("sheet"+sheetIndex);
		// 创建第1行
		Row row = sheet.createRow(0);
		// 设置第1行的数据
		W_PoiOffice.setTitle(row, vos);
		
		// 第二行开始，填充数据
		int j = 1;
		for (int i = 0; i < list.size(); i++) {
			row = sheet.createRow(j++);
			W_PoiOffice.setData(wb, row, vos, list.get(i));
		}
		return true;
	}
	/**
	 * 将集合中的数据，写到excel文档中
	 * @param filePath excel文档的路径
	 * @param tablename excel工作表名称
	 * @param list 数据集合
	 * @return boolean 真假
	 */
	public static boolean writeToExcel_xls(String filePath, String tablename, List<?> list) {
		OutputStream stream = null;
		try {
			stream=new FileOutputStream(filePath);
			return writeToExcel_xls(stream,tablename, list);
		}catch (Exception e) {
			return false;
		}finally {
			if(null!=stream) {
				try {
					stream.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}
	
	/**
	 * 将集合中的数据，通过输出流，写到excel文档中
	 * @param stream 输出流
	 * @param tablename excel工作表名称
	 * @param list 数据集合
	 * @return boolean 真假
	 */
	public static boolean writeToExcel_xls(OutputStream stream, String tablename, List<?> list) {
		if (null == list || list.isEmpty()) {
			return false;
		}
		Class<?> c = list.get(0).getClass();
		JWEOfficeVO[] vos=null;
		try {
			vos = JWEOfficeVO.getJWEOfficeVO(c);
		} catch (Exception e1) {
			e1.printStackTrace();
			return false;
		}
		try (HSSFWorkbook wb = new HSSFWorkbook()) {
			// 创建一张表格
			HSSFSheet sheet = wb.createSheet(null == tablename || tablename.isEmpty() ? "sheet1" : tablename);
			// 创建第1行
			HSSFRow row = sheet.createRow(0);
			// 设置第1行的数据
			W_PoiOffice.setTitle(row, vos);

			// 第二行开始，填充数据
			int j = 1;
			for (int i = 0; i < list.size(); i++) {
				row = sheet.createRow(j++);
				W_PoiOffice.setData(wb, row, vos, list.get(i));
			}
			wb.write(stream);
			return true;
		}catch (Exception e) {
			return false;
		}
	}
	
	/**
	 * 将集合中的数据，写到excel文档中
	 * @param filePath excel文档的路径
	 * @param tablename excel工作表名称
	 * @param list 数据集合
	 * @return boolean
	 */
	public static boolean writeToExcel_xlsx(String filePath, String tablename, List<?> list) {
		OutputStream stream = null;
		try {
			stream=new FileOutputStream(filePath);
			return writeToExcel_xlsx(stream,tablename, list);
		}catch (Exception e) {  
			return false;
		}finally {
			if(null!=stream) {
				try {
					stream.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}
	
	/**
	 * 将集合中的数据，通过输出流，写到excel文档中
	 * @param stream 输出流
	 * @param tablename excel工作表名称
	 * @param list 数据集合
	 * @return boolean
	 */
	public static boolean writeToExcel_xlsx(OutputStream stream, String tablename, List<?> list) {
		if (null == list || list.isEmpty()) {
			return false;
		}
		Class<?> c = list.get(0).getClass();
		JWEOfficeVO[] vos=null;
		try {
			vos = JWEOfficeVO.getJWEOfficeVO(c);
		} catch (Exception e1) {
			e1.printStackTrace();
			return false;
		}
		try (XSSFWorkbook wb = new XSSFWorkbook()) {
			// 创建一张表格
			XSSFSheet sheet = wb.createSheet(null == tablename || tablename.isEmpty() ? "sheet1" : tablename);
			// 创建第1行
			XSSFRow row = sheet.createRow(0);
			// 设置第1行的数据
			W_PoiOffice.setTitle(row, vos);

			// 第二行开始，填充数据
			int j = 1;
			for (int i = 0; i < list.size(); i++) {
				row = sheet.createRow(j++);
				W_PoiOffice.setData(wb, row, vos, list.get(i));
			}
			wb.write(stream);
			return true;
		}catch (Exception e) {
			return false;
		}
	}
}
