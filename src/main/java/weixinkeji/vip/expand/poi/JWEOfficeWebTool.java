package weixinkeji.vip.expand.poi;

import java.io.IOException;
import java.util.List;
import java.util.Map;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServletResponse;

public class JWEOfficeWebTool {

	/**
	 * 把集合的数据，写成excel文档，并通过输出流，输出给客户端
	 * 
	 * @param resp           HttpServletResponse
	 * @param list           用户的数据
	 * @param fileName       下载显示的文件名（autoSetHeader=true时才起作用）
	 * @param excelSheetName excel文档——工作表的名
	 * @param autoSetHeader  是否自动设置报文头:关于下载文件的
	 * @return boolean
	 * @throws ServletException 异常
	 * @throws IOException      异常
	 */
	public static boolean downloadExcelFile(HttpServletResponse resp, List<?> list, String fileName,
			String excelSheetName,boolean autoSetHeader) throws ServletException, IOException {
		if(autoSetHeader) {
			setFileContentHeader(resp, fileName);
		}
		return fileName.endsWith("xls") ? JWEOfficeW.writeToExcel_xls(resp.getOutputStream(), excelSheetName, list)
				: JWEOfficeW.writeToExcel_xlsx(resp.getOutputStream(), excelSheetName, list);
	}

	/**
	 * 把集合的数据，写成excel文档，并通过输出流，输出给客户端(会自动调用关闭）
	 * 
	 * @param resp     HttpServletResponse
	 * @param data     用户的数据 key是 excel文档——工作表的名；value是 excel文档——工作表的数据
	 * @param fileName 下载显示的文件名（autoSetHeader=true时才起作用）
	 * @param autoSetHeader  是否自动设置报文头:关于下载文件的
	 * @return boolean
	 * @throws ServletException 异常
	 * @throws IOException      异常
	 */
	public static <T>boolean downloadExcelFile_autoClose(HttpServletResponse resp, Map<String, List<T>> data, String fileName,boolean autoSetHeader)
			throws ServletException, IOException {
		if(autoSetHeader) {
			setFileContentHeader(resp, fileName);
		}
		JWEOfficeW obj = new JWEOfficeW(resp.getOutputStream(),
				fileName.endsWith(".xls") ? JWEOfficeEnum.xls : JWEOfficeEnum.xlsx);
		for (Map.Entry<String, List<T>> kv : data.entrySet()) {
			try {
				if(!obj.addToExcel(kv.getKey(), kv.getValue())) {
					obj.writeAndAutoCloseIO();
					return false;
				}
			} catch (Exception e) {
				e.printStackTrace();
				return false;
			}
		}
		obj.writeAndAutoCloseIO();
		return true;
	}
	
	
	/**
	 * 把集合的数据，写成excel文档，并通过输出流，输出给客户端
	 * 
	 * @param resp     HttpServletResponse
	 * @param data     用户的数据 key是 excel文档——工作表的名；value是 excel文档——工作表的数据
	 * @param fileName 下载显示的文件名（autoSetHeader=true时才起作用）
	 * @param autoSetHeader  是否自动设置报文头:关于下载文件的
	 * @return boolean
	 * @throws ServletException 异常
	 * @throws IOException      异常
	 */
	public static <T>boolean downloadExcelFile(HttpServletResponse resp, Map<String, List<T>> data, String fileName,boolean autoSetHeader)
			throws ServletException, IOException {
		if(autoSetHeader) {
			setFileContentHeader(resp, fileName);
		}
		JWEOfficeW obj = new JWEOfficeW(resp.getOutputStream(),
				fileName.endsWith(".xls") ? JWEOfficeEnum.xls : JWEOfficeEnum.xlsx);
		for (Map.Entry<String, List<T>> kv : data.entrySet()) {
			try {
				if(!obj.addToExcel(kv.getKey(), kv.getValue())) {
					return false;
				}
			} catch (Exception e) {
				e.printStackTrace();
				return false;
			}
		}
		return true;
	}
	
	
	
	// 设置下载时，文件名的显示 及编号 处理
	private static void setFileContentHeader(HttpServletResponse resp, String filename) throws IOException {
		resp.setHeader("Content-Disposition",
				"attachment;filename=" + new String(filename.getBytes("utf-8"), "ISO8859-1"));
	}
}
