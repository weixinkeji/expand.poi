package weixinkeji.vip.expand.poi;

import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

public class JWEOfficeVO {
	private static final Map<Class<?>, JWEOfficeVO[]> map = new HashMap<>();

	/**
	 * 取得注解在类上的数据，转化成对象（一个属性上的注解，等于一个对象。多个属性，多个对象。返回对象数组）
	 * 
	 * @param c 目标类
	 * @return JWEOfficeVO[] 对象数组
	 * @throws Exception 异常
	 */
	synchronized public static JWEOfficeVO[] getJWEOfficeVO(Class<?> c) throws Exception {
		JWEOfficeVO[] vo = map.get(c);
		if (null == vo) {
			vo = new JWEOfficeVOTool().getJWEOfficeVO(c);
			map.put(c, vo);
		}
		return vo;
	}
	/**
	 * 利用反射，给对象附值
	 * @param obj 对象
	 * @param value 值
	 * @throws Exception 异常
	 */
	public void setValue(Object obj,Object value) throws Exception{
		if(null==value) {
			return;
		}
		String str;
		switch (this.valueType) {
		case "String":
			this.getField().set(obj, value);
			return;
		case "boolean":
		case "Boolean":
			if(value.getClass().getSimpleName().equalsIgnoreCase("boolean")) {
				this.getField().set(obj, value);
			}else {
				str=String.valueOf(value);
				if(str.isEmpty()) {
					 return;
				 }
				this.getField().set(obj, Boolean.parseBoolean(str));
			}
			return;
		case "short":
		case "Short":
			 str=String.valueOf(value);
			 if(str.isEmpty()) {
				 return;
			 }
			 if(str.contains(".")) {
				 str=str.substring(0,str.indexOf("."));
			 }
			this.getField().set(obj, Short.parseShort(str));
			return;
		case "int":
		case "Integer":
			str=String.valueOf(value);
			if(str.isEmpty()) {
				 return;
			 }
			 if(str.contains(".")) {
				 str=str.substring(0,str.indexOf("."));
			 }
			this.getField().set(obj, Integer.parseInt(str));
			return;
		case "long":
		case "Long":
			str=String.valueOf(value);
			if(str.isEmpty()) {
				 return;
			 }
			 if(str.contains(".")) {
				 str=str.substring(0,str.indexOf("."));
			 }
			this.getField().set(obj, Long.parseLong(str));
			return;
		case "float":
		case "Float":
			str=String.valueOf(value);
			if(str.isEmpty()) {
				 return;
			 }
			this.getField().set(obj, Long.parseLong(str));
			return;
		case "double":
		case "Double":
			if(value.getClass().getSimpleName().equalsIgnoreCase("double")) {
				this.getField().set(obj, value);
			}else {
				str=String.valueOf(value);
				if(str.isEmpty()) {
					 return;
				 }
				this.getField().set(obj, Double.parseDouble(str));
			}
			return;
		case "Date":
			if(value.getClass().getSimpleName().equalsIgnoreCase("double")){
				this.getField().set(obj,HSSFDateUtil.getJavaDate((double)value));
			}else if(value.getClass().getSimpleName().equalsIgnoreCase("Date")) {
				this.getField().set(obj, value);
			} else {
				str=String.valueOf(value);
				if(str.isEmpty()) {
					 return;
				 }
				SimpleDateFormat sf=new SimpleDateFormat(this.dateFormat);
				this.getField().set(obj, sf.parse(str));
			}
		}
	}

	private int sort;// 第几列
	private String title;// excel里的标题名
	private String valueType; // 值类型
//	private Object value;// 值
	private String dateFormat; // 时间格式
	private Field field;// 对应类的属性对象

	public Field getField() {
		return field;
	}

	public void setField(Field field) {
		this.field = field;
	}

	public static Map<Class<?>, JWEOfficeVO[]> getMap() {
		return map;
	}

	public int getSort() {
		return sort;
	}

	public void setSort(int sort) {
		this.sort = sort;
	}

	public String getTitle() {
		return title;
	}

	public void setTitle(String title) {
		this.title = title;
	}

	public String getValueType() {
		return valueType;
	}

	public void setValueType(String valueType) {
		this.valueType = valueType;
	}

	public String getDateFormat() {
		return dateFormat;
	}

	public void setDateFormat(String dateFormat) {
		this.dateFormat = dateFormat;
	}
}

class JWEOfficeVOTool {

	public JWEOfficeVO[] getJWEOfficeVO(Class<?> c) throws Exception {
		Field[] fs = c.getDeclaredFields();
		Set<JWEOfficeVO> set = new HashSet<>();
		JWEOfficeVO vo;
		for (Field f : fs) {
			f.setAccessible(true);
			vo = this.getJWEOfficeVO(f);
			if (null != vo) {
				set.add(this.getJWEOfficeVO(f));
			}
		}
		JWEOfficeVO[] vos = new JWEOfficeVO[set.size()];
		set.toArray(vos);
		this.sort(vos);
		return vos;
	}

	private JWEOfficeVO getJWEOfficeVO(Field f) throws Exception {
		JWEOffice ann = f.getAnnotation(JWEOffice.class);
		if (null == ann) {
			return null;
		}
		JWEOfficeVO obj = new JWEOfficeVO();
		// 设置此标题在第几列
		obj.setSort(ann.sort());
		// 设置标题
		if (null != ann.title() && ann.title().length() > 0) {
			obj.setTitle(ann.title());
		} else if (null != ann.value() && ann.value().length() > 0) {
			obj.setTitle(ann.value());
		} else {
			obj.setTitle(f.getName());
		}
		// 设置值类型
		obj.setValueType(f.getType().getSimpleName());
		// 设置时间格式
		obj.setDateFormat(ann.dateformat());
		// 设置归属属性
		obj.setField(f);
		return obj;
	}
	private void sort(JWEOfficeVO[] vo) {
		JWEOfficeVO jv;
		for (int i = 0; i < vo.length - 1; i++) {
			for (int j = 0; j < vo.length - 1 - i; j++) {
				if (vo[j].getSort() > vo[j + 1].getSort()) {
					jv = vo[j];
					vo[j] = vo[j + 1];
					vo[j + 1] = jv;
				}
			}
		}
	}
}

class R_PoiOffice<T> {
	private Map<Integer, JWEOfficeVO> cellIndexMapJWEOfficeVO = new HashMap<Integer, JWEOfficeVO>();
	private JWEOfficeVO[] vo;
	private Class<T> c;

	public R_PoiOffice(Class<T> t) throws Exception {
		vo = JWEOfficeVO.getJWEOfficeVO(t);
		c = t;
	}

	// 初始化 第几列对应哪个 JWEOfficeVO
	public void init_cellMapJWEOfficeVO(Row row) {
		Cell cell;
		for (int i = 0; i < vo.length; i++) {
			cell = row.getCell(i);
			if (null != cell) {
				cellIndexMapJWEOfficeVO.put(i, getJWEOfficeVO(cell.getStringCellValue()));
			}
		}
	}
	private JWEOfficeVO getJWEOfficeVO(String title) {
		for (JWEOfficeVO v : vo) {
			if (v.getTitle().equalsIgnoreCase(title)) {
				return v;
			}
		}
		return null;
	}
	
	
	public T readRow(Row row) throws Exception {
		T obj = this.c.getConstructor().newInstance();
		Cell cell;
		JWEOfficeVO vo;
		for (Map.Entry<Integer, JWEOfficeVO> kv : cellIndexMapJWEOfficeVO.entrySet()) {
			cell = row.getCell(kv.getKey());
			vo = kv.getValue();
			if(null!=cell) {
				this.setValueByCell(vo, obj, cell);
			}
		}
		return obj;
	}

	private void setValueByCell(JWEOfficeVO vo,Object obj,Cell cell) throws Exception {
		Object value;
		switch (cell.getCellType()) {
		case FORMULA:
			value =cell.getCellFormula();
			break;
		case NUMERIC:
			value = cell.getNumericCellValue();
			break;
		case STRING:
			value =cell.getStringCellValue();
			break;
		case BLANK:
			value = "";
			break;
		case BOOLEAN:
			value =cell.getBooleanCellValue();
			break;
		case ERROR:
			value =  cell.getErrorCellValue();
			break;
		default:
			value =cell.getCellType();
		}
		vo.setValue(obj, value);
	}
}


class W_PoiOffice {
	/**
	 * 设置第一行的值（作用标题）
	 * 
	 * @param row
	 * @param vo
	 */
	public static void setTitle(Row row, JWEOfficeVO[] vos) {
		int i = 0;
		for (JWEOfficeVO vo : vos) {
			row.createCell(i++).setCellValue(vo.getTitle());
		}
	}

	/**
	 * 设置第x行的数据值
	 * 
	 * @param row POI框架-行对象
	 * @param vos JWEOffice对象模型 的集合
	 * @param obj 用户的数据（用于填充表格里的单元格）
	 * @throws Exception
	 */
	public static void setData(Workbook wb, Row row, JWEOfficeVO[] vos, Object obj) throws Exception {
		JWEOfficeVO vo;
		Cell cell;
		Object value;
		for (int i = 0; i < vos.length; i++) {
			cell = row.createCell(i);
			vo = vos[i];
			switch (vos[i].getValueType()) {
			case "String":
				if(null==(value=vo.getField().get(obj))) {
					continue;
				}
				cell.setCellValue(value.toString());
				continue;
			case "boolean":
			case "Boolean":
				if(null==(value=vo.getField().get(obj))) {
					continue;
				}
				cell.setCellValue(Boolean.parseBoolean(value.toString()));
				continue;
			case "short":
			case "Short":
			case "int":
			case "Integer":
			case "long":
			case "Long":
			case "float":
			case "Float":
			case "double":
			case "Double":
				if(null==(value=vo.getField().get(obj))) {
					continue;
				}
				cell.setCellValue(Double.parseDouble(value.toString()));
				continue;
			case "Date":
				if(null==(value=vo.getField().get(obj))) {
					continue;
				}
				setDate(wb, cell, (Date) value, vo.getDateFormat());
			}
		}
	}

	/**
	 * 给单元格cell，附上时间值
	 * 
	 * @param wb         Workbook对象
	 * @param cell       单元格 对象
	 * @param date       用户时间
	 * @param dateFormat 时间格式
	 */
	public static void setDate(Workbook wb, Cell cell, Date date, String dateFormat) {
		CellStyle style = wb.createCellStyle();
		style.setDataFormat(wb.getCreationHelper().createDataFormat().getFormat(dateFormat));
		cell.setCellValue(new Date());
		cell.setCellStyle(style);
	}
	
	/**
	 * 给单元格cell，附上 超连接值
	 * 
	 * @param cell
	 * @param url
	 * @param showUrlName
	 */
	public static void setHyperlink(Cell cell, String url, String showUrlName) {
		cell.setCellFormula("HYPERLINK(\"" + url + "\",\"" + showUrlName + "\")");
	}
}


