package weixinkeji.vip.expand.poi;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class JWEOfficeR {
	private Workbook wb;
	/**
	 * 构造方法
	 * @param fis InputStream 输入流
	 * @param xls_xlsx JWEOfficeEnum  元素 指明是xlx还是xlsx格式
	 * @throws IOException io流异常
	 */
	public JWEOfficeR(InputStream fis,JWEOfficeEnum xls_xlsx) throws IOException{
		ZipSecureFile.setMinInflateRatio(-1.0d);
		this.wb = xls_xlsx==JWEOfficeEnum.xls?
				new HSSFWorkbook(fis)
				:new XSSFWorkbook(fis)
				;
	}
	/**
	 * 取得 Workbook 接口的实例。
	 * @return Workbook
	 */
	public Workbook getWorkbook() {
		return this.wb;
	}
	/**
	 * 构造方法
	 * @param filePath String 文件路径
	 * @param xls_xlsx JWEOfficeEnum  元素 指明是xlx还是xlsx格式
	 * @throws IOException io流异常
	 */
	public JWEOfficeR(String filePath,JWEOfficeEnum xls_xlsx) throws IOException{
		InputStream fis = null;
		try {
			ZipSecureFile.setMinInflateRatio(-1.0d);
			fis=new FileInputStream(filePath);
			this.wb = xls_xlsx==JWEOfficeEnum.xls?
					new HSSFWorkbook(fis)
					:new XSSFWorkbook(fis)
					;
		}catch (Exception e) {
			throw e;
		}finally {
			if(null!=fis) {
				fis.close();
			}
		}
	}
	
	/**
	 * 通过输入流，读取excel文档的内容 到 集合中
	 * @param <T> 相关的特征类
	 * @param c   相关的特征类
	 * @param sheetName excel工作表名称
	 * @return		List 集合
	 * @throws Exception 异常
	 */
	public <T>List<T> readExcel(Class<T> c,String sheetName) throws Exception {
		R_PoiOffice<T> robj=new R_PoiOffice<T>(c);
			Sheet sheet = this.wb.getSheet(sheetName);
			int rows = sheet.getPhysicalNumberOfRows();
			if(rows<1) {
				return null;
			}
			List<T> list=new ArrayList<>();
			robj.init_cellMapJWEOfficeVO(sheet.getRow(0));
			
			for (int r =1; r < rows; r++) {
				Row row = sheet.getRow(r);
				if (row == null) {
					continue;
				}
				list.add(robj.readRow(row));
			}
			return list;
	}
	/**
	 * 通过输入流，读取excel文档的内容 到 集合中
	 * @param <T> 相关的特征类
	 * @param c   相关的特征类
	 * @param sheetIndex 第几个excel工作表
	 * @return		List 集合
	 * @throws Exception 异常
	 */
	public <T>List<T> readExcel(Class<T> c,int sheetIndex) throws Exception {
		R_PoiOffice<T> robj=new R_PoiOffice<T>(c);
			Sheet sheet = this.wb.getSheetAt(sheetIndex);
			int rows = sheet.getPhysicalNumberOfRows();
			if(rows<1) {
				return null;
			}
			List<T> list=new ArrayList<>();
			robj.init_cellMapJWEOfficeVO(sheet.getRow(0));
			
			for (int r =1; r < rows; r++) {
				Row row = sheet.getRow(r);
				if (row == null) {
					continue;
				}
				list.add(robj.readRow(row));
			}
			return list;
	}
	
	 /**
	  * 读取excel文档的内容 到 集合中
	 * @param <T> 相关的特征类
	 * @param c   相关的特征类
	 * @param sheetName excel工作表名称
	 * @param filePath excel文档的路径
	 * @return List 集合
	 * @throws Exception 异常
	 */
	public static <T>List<T> readExcel_xls(Class<T> c,String sheetName, String filePath) throws Exception {
		InputStream fis = null;
		try {
			fis=new FileInputStream(filePath);
			return readExcel_xls(c,sheetName, fis);
		}catch (Exception e) {
			throw e;
		}finally {
			if(null!=fis) {
				fis.close();
			}
		}
	}
	
	/**
	 * 通过输入流，读取excel文档的内容 到 集合中
	 * @param <T> 相关的特征类
	 * @param c   相关的特征类
	 * @param sheetName excel工作表名称
	 * @param fis	输入流
	 * @return		List 集合
	 * @throws Exception 异常
	 */
	public static <T>List<T> readExcel_xls(Class<T> c,String sheetName, InputStream fis) throws Exception {
		R_PoiOffice<T> robj=new R_PoiOffice<T>(c);
		ZipSecureFile.setMinInflateRatio(-1.0d);
		try (HSSFWorkbook wb = new HSSFWorkbook(fis)) {
			HSSFSheet sheet = wb.getSheet(sheetName);
			int rows = sheet.getPhysicalNumberOfRows();
			if(rows<1) {
				return null;
			}
			List<T> list=new ArrayList<>();
			robj.init_cellMapJWEOfficeVO(sheet.getRow(0));
			
			for (int r =1; r < rows; r++) {
				HSSFRow row = sheet.getRow(r);
				if (row == null) {
					continue;
				}
				list.add(robj.readRow(row));
			}
			return list;
		}
	}

	 /**
	  * 读取excel文档的内容 到 集合中
	 * @param <T> 相关的特征类
	 * @param c   相关的特征类
	 * @param sheetName excel工作表名称
	 * @param filePath excel文档的路径
	 * @return List 集合
	 * @throws Exception 异常
	 */
	public static <T>List<T> readExcel_xlsx(Class<T> c,String sheetName, String filePath) throws Exception {
		InputStream fis = null;
		try {
			fis=new FileInputStream(filePath);
			return readExcel_xlsx(c,sheetName, fis);
		}catch (Exception e) {
			throw e;
		}finally {
			if(null!=fis) {
				fis.close();
			}
		}
	}
	
	/**
	 * 通过输入流，读取excel文档的内容 到 集合中
	 * @param <T> 相关的特征类
	 * @param c   相关的特征类
	 * @param sheetName excel工作表名称
	 * @param fis	输入流
	 * @return		List
	 * @throws Exception 异常
	 */
	public static <T>List<T> readExcel_xlsx(Class<T> c,String sheetName, InputStream fis) throws Exception {
		R_PoiOffice<T> robj=new R_PoiOffice<T>(c);
		ZipSecureFile.setMinInflateRatio(-1.0d);
		try (XSSFWorkbook wb = new XSSFWorkbook(fis)) {
			XSSFSheet sheet = wb.getSheet(sheetName);
			int rows = sheet.getPhysicalNumberOfRows();
			if(rows<1) {
				return null;
			}
			List<T> list=new ArrayList<>();
			robj.init_cellMapJWEOfficeVO(sheet.getRow(0));
			
			for (int r =1; r < rows; r++) {
				Row row = sheet.getRow(r);
				if (row == null) {
					continue;
				}
				list.add(robj.readRow(row));
			}
			return list;
		}
	}
}
