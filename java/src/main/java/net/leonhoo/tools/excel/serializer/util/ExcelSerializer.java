package net.leonhoo.tools.excel.serializer.util;

import java.io.File;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.Type;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Colour;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import net.leonhoo.tools.excel.serializer.adapter.IColorPicker;
import net.leonhoo.tools.excel.serializer.adapter.IValueConverter;
import net.leonhoo.tools.excel.serializer.annotation.Align;
import net.leonhoo.tools.excel.serializer.annotation.Align.EnumCellAlign;
import net.leonhoo.tools.excel.serializer.annotation.Color;
import net.leonhoo.tools.excel.serializer.annotation.Index;
import net.leonhoo.tools.excel.serializer.annotation.Indexs;
import net.leonhoo.tools.excel.serializer.annotation.Width;
import net.leonhoo.tools.excel.serializer.annotation.Display;
import net.leonhoo.tools.excel.serializer.annotation.Converter;

/**
 * Excel自动序列化类
 * 
 * @author leon
 *
 * @param <T>
 */
public class ExcelSerializer<T> {
	// Export
	private Map<Field, String> mapDisplayName = new HashMap<Field, String>();// <属性名,显示名称>
	private Map<Field, jxl.format.Alignment> mapAlign = new HashMap<Field, jxl.format.Alignment>();// <属性名,显示位置>
	private Map<Field, Integer> mapWidth = new HashMap<Field, Integer>();// <属性名,显示宽度>
	private Map<Field, IColorPicker> mapColorConverter = new HashMap<Field, IColorPicker>();// <属性名,颜色转换器>
	// Import
	private Map<Field, Integer> mapColumnIndex = new HashMap<Field, Integer>();// <属性名,列序号>
	private Map<Field, Integer[]> mapColumnIndexs /* <属性名,列序号数组> */
	= new HashMap<Field, Integer[]>();// 用于转换二维矩阵
	private Map<Field, IValueConverter> mapValueConverter = new HashMap<Field, IValueConverter>();// <属性名,值转换器实例>

	// Excel文件对应类
	private Class<T> clazz;
	private Field[] fields;
	private String clazzDisplayName;// 类显示名称,对应Excel的Sheet的名称
	private jxl.format.Alignment defaultCellAlign = Alignment.CENTRE;// 默认输出ExcelCell位置
	private int defaultCellWidth = 10;
	private int defaultCellHeight = 370;
	private boolean isIgnoreExcelHeader = true; // 是否忽略Excel的表头,即不读取第一行
	// 循环次数: 用于循环获取一行中的数据,赋予集合中的子对象
	private int loopTime = 0;

	/**
	 * 设置默认单元格位置
	 * 
	 * @param defaultCellAlign
	 */
	public void setDefaultCellAlign(jxl.format.Alignment defaultCellAlign) {
		this.defaultCellAlign = defaultCellAlign;
	}

	/**
	 * 设置默认列宽
	 * 
	 * @param defaultCellWidth
	 */
	public void setDefaultCellWidth(int defaultCellWidth) {
		this.defaultCellWidth = defaultCellWidth;
	}

	/**
	 * 设置默认行高
	 * 
	 * @return
	 */
	public void setDefaultCellHeight(int defaultCellHeight) {
		this.defaultCellHeight = defaultCellHeight;
	}

	/**
	 * 是否忽略Excel表头(忽略Excel表第一行)
	 * 
	 * @param isIgnore
	 */
	public void setIgnoreExcelHeader(boolean isIgnore) {
		this.isIgnoreExcelHeader = isIgnore;
	}

	public ExcelSerializer(Class<T> clazz) throws Exception {
		this.clazz = clazz;
		this.fields = clazz.getDeclaredFields();
		this.clazzDisplayName = this.clazz.getSimpleName();
		if (clazz.isAnnotationPresent(Display.class)) {
			this.clazzDisplayName = clazz.getAnnotation(Display.class).value();
		}
		getExcelColumnDict(clazz);
	}

	public List<T> load(String uri) throws Exception {
		List<T> result = null;
		File f = new File(uri);// Excel文件
		if (f.exists()) {
			Sheet sheet = getSheet(f);
			if (sheet == null)
				throw new Exception("表不存在,请检查表名");

			if (sheet.getRows() > 0) {
				int limitEmptyRow = 5; // 最大允许5个连续空行（超出5行则不循环下面的数据了）
				int emptyRow = 0; // 记录连续空行的个数

				result = new ArrayList<T>();
				T t = null;
				int start = this.isIgnoreExcelHeader ? 1 : 0;// 表头的目录不需要，从1开始
				for (int i = start; i < sheet.getRows(); i++) { // 行数
					if (emptyRow >= limitEmptyRow)
						break; // 最大允许连续空行
					if (sheet.getCell(0, i).getContents() == "") {
						emptyRow++;
						continue;
					}
					t = this.clazz.newInstance();
					// 开始赋值
					for (Field field : fields) {
						if (mapColumnIndex.containsKey(field)) {
							int index = mapColumnIndex.get(field);
							// 读取Excel指定index列
							String cellValue = sheet.getCell(index, i).getContents();
							Object objValue = null;
							if (mapValueConverter.containsKey(field)) {
								IValueConverter converter = mapValueConverter.get(field);
								objValue = converter.deserialize(cellValue, field.getType(), null);
							} else {
								objValue = changeType(cellValue, field.getType());
							}
							String name = field.getName();
							String methodName = String.format("set%s%s", name.substring(0, 1).toUpperCase(),
									name.substring(1));
							Method method = clazz.getDeclaredMethod(methodName, new Class[] { field.getType() });
							method.invoke(t, objValue);
						}
					}
					result.add(t);
				}
			}
		}
		return result;
	}

	public Map<Integer, List<T>> loadList(String uri) throws Exception {
		Map<Integer, List<T>> result = null;
		File f = new File(uri);// Excel文件
		if (f.exists()) {
			Sheet sheet = getSheet(f);
			if (sheet == null)
				throw new Exception("表不存在,请检查表名");

			if (sheet.getRows() > 0) {
				int limitEmptyRow = 5; // 最大允许5个连续空行（超出5行则不循环下面的数据了）
				int emptyRow = 0; // 记录连续空行的个数

				result = new HashMap<Integer, List<T>>();
				List<T> ts = null;
				int start = this.isIgnoreExcelHeader ? 1 : 0;// 表头的目录不需要，从1开始
				for (int i = start; i < sheet.getRows(); i++) { // 行数
					if (emptyRow >= limitEmptyRow)
						break; // 最大允许连续空行
					if (sheet.getCell(0, i).getContents() == "") {
						emptyRow++;
						continue;
					}
					ts = new ArrayList<T>();
					// 预先生成子对象
					for (int j = 0; j < this.loopTime; j++) {
						ts.add(this.clazz.newInstance());
					}
					result.put((i - start), ts);
					// 开始赋值
					for (Field field : fields) {
						if (mapColumnIndexs.containsKey(field)) {
							Integer[] indexs = mapColumnIndexs.get(field);
							for (int j = 0; j < indexs.length; j++) {
								T t = ts.get(j);
								// 读取Excel指定index列
								String cellValue = sheet.getCell(indexs[j], i).getContents();
								Object objValue = null;
								if (mapValueConverter.containsKey(field)) {
									IValueConverter converter = mapValueConverter.get(field);
									objValue = converter.deserialize(cellValue, field.getType(), null);
								} else {
									objValue = changeType(cellValue, field.getType());
								}
								String name = field.getName();
								String methodName = String.format("set%s%s", name.substring(0, 1).toUpperCase(),
										name.substring(1));
								Method method = clazz.getDeclaredMethod(methodName, new Class[] { field.getType() });
								method.invoke(t, objValue);
							}
						}
					}
					result.put(i, ts);
				}
			}
		}
		return result;
	}

	/**
	 * 重新生成新的Excel文件
	 * 
	 * @param uri
	 *            Excel文件路径
	 * @param data
	 *            需要保存的数据
	 * @param isAppendData
	 *            是否追加数据
	 * @throws Exception
	 */
	public void save(String uri, List<T> data, boolean isAppendData) throws Exception {
		save(uri, data, 0, isAppendData);
	}

	/**
	 * 在原有Excel中追加表
	 * 
	 * @param uri
	 *            Excel文件路径
	 * @param data
	 *            需要保存的数据
	 * @param appendSheetIndex
	 *            追加表的位置
	 * @param isAppendData
	 *            是否追加数据
	 * @throws Exception
	 */
	public void save(String uri, List<T> data, int appendSheetIndex, boolean isAppendData) throws Exception {
		File f = new File(uri);// Excel文件
		WritableWorkbook workbook = null;
		Workbook wb = null;
		if (f.exists()) {
			try {
				wb = Workbook.getWorkbook(f);
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		if (wb == null || !isAppendData) {
			workbook = Workbook.createWorkbook(f);
		} else {
			workbook = Workbook.createWorkbook(f, wb);
		}

		// 创建Excel工作表 指定名称和位置
		WritableSheet sheet = workbook.getSheet(this.clazzDisplayName);
		if (sheet == null || !isAppendData) {
			sheet = workbook.createSheet(this.clazzDisplayName, appendSheetIndex);
		}
		// **************往工作表中添加数据*****************
		setHeader(sheet);
		setBody(sheet, data);
		// **************添加数据结束*****************
		// 写入工作表完毕，关闭流
		workbook.write();
		workbook.close();
	}

	private void setHeader(WritableSheet sheet) throws Exception {
		// 设置格式
		WritableFont titleFont = new WritableFont(WritableFont.createFont("微软雅黑"), 11, WritableFont.NO_BOLD);
		WritableCellFormat titleFormat = new WritableCellFormat(titleFont);
		titleFormat.setAlignment(jxl.format.Alignment.CENTRE);
		for (Field field : fields) {
			if (mapColumnIndex.containsKey(field)) {
				int index = mapColumnIndex.get(field);
				// 显示的名称
				String displayName = field.getName();
				if (mapDisplayName.containsKey(field)) {
					displayName = mapDisplayName.get(field);
				}
				// 显示的宽度
				int cellWidth = this.defaultCellWidth;
				if (mapWidth.containsKey(field)) {
					cellWidth = mapWidth.get(field);
				}
				// 生成标题栏
				sheet.addCell(new Label(index, 0, displayName, titleFormat));
				sheet.setColumnView(index, cellWidth);
				sheet.setRowView(0, this.defaultCellHeight, false); // 设置行高
			}
		}
	}

	private void setBody(WritableSheet sheet, List<T> data) throws Exception {
		WritableFont contentFont = new WritableFont(WritableFont.createFont("楷体 _GB2312"), 11, WritableFont.NO_BOLD);
		WritableCellFormat contentFormat = null;
		for (int i = 0; i < data.size(); i++) {
			T t = data.get(i);
			for (Field field : fields) {
				if (mapColumnIndex.containsKey(field)) {
					// 显示的列的序号
					int index = mapColumnIndex.get(field);
					// 获取显示的值
					String name = field.getName();
					String methodName = String.format("get%s%s", name.substring(0, 1).toUpperCase(), name.substring(1));
					Method method = clazz.getDeclaredMethod(methodName);
					Object objValue = method.invoke(t);// 获取值
					// 获取Cell显示位置
					Alignment currentAlign = this.defaultCellAlign;
					if (mapAlign.containsKey(field)) {
						currentAlign = mapAlign.get(field);
					}
					contentFormat = new WritableCellFormat(contentFont);
					contentFormat.setAlignment(currentAlign);
					// 根据值判断,设置颜色
					if (mapColorConverter.containsKey(field)) {
						String colorString = mapColorConverter.get(field).get(objValue, field.getType(), null);
						if (colorString != null && colorString.length() > 0) {
							contentFormat.setBackground(getNearestColour(colorString));
						}
					}
					// 值转换
					if (mapValueConverter.containsKey(field)) {
						objValue = mapValueConverter.get(field).serialize(objValue, String.class, null);
					}
					// 赋值
					sheet.addCell(new Label(index, i + 1, objValue.toString(), contentFormat));
				}
			}
			sheet.setRowView(0, this.defaultCellHeight, false); // 设置行高
		}
	}

	/**
	 * 从Excel中获取对应表
	 * 
	 * @param f
	 * @return
	 * @throws Exception
	 */
	private Sheet getSheet(File f) throws Exception {
		Workbook rwb = Workbook.getWorkbook(f); // 获取Excel文件对象
		Sheet sheet = null;
		if (this.clazzDisplayName != null && this.clazzDisplayName.length() > 0) {
			sheet = rwb.getSheet(this.clazzDisplayName); // 指定工作表
		}
		if (sheet == null) {
			sheet = rwb.getSheet(0); // 获取文件的指定工作表 默认的第一个
		}
		return sheet;
	}

	/**
	 * 将值转换成指定类型
	 * 
	 * @param v
	 * @param type
	 * @return
	 * @throws Exception
	 */
	private Object changeType(String v, Type type) throws Exception {
		String tmpValue = (v + "").trim();
		if (tmpValue == null || tmpValue.length() < 1) {
			if (type == String.class) {
				return "";
			} else {
				return null;
			}
		}

		Object result = tmpValue;
		if (type == boolean.class || type == Boolean.class) {
			result = ("true".equalsIgnoreCase(tmpValue) || "1".equals(tmpValue));
		} else if (type == int.class || type == Integer.class) {
			result = Integer.parseInt(tmpValue);
		} else if (type == double.class || type == Double.class) {
			result = Double.parseDouble(tmpValue);
		} else if (type == float.class || type == Float.class) {
			result = Float.parseFloat(tmpValue);
		} else if (type == Date.class) {
			SimpleDateFormat sdf = new SimpleDateFormat();
			result = sdf.parse(tmpValue);
		} else if (type == short.class || type == Short.class) {
			result = Short.parseShort(tmpValue);
		} else if (type == long.class || type == Long.class) {
			result = Long.parseLong(tmpValue);
		} else if (type == byte.class || type == Byte.class) {
			result = Byte.parseByte(tmpValue);
		} else if (type == char.class) {
			result = tmpValue.charAt(0);
		}
		return result;
	}

	/**
	 * 获取该类中所有注解,并放入与属性对应的Map
	 * 
	 * @param clazz
	 * @throws Exception
	 */
	private void getExcelColumnDict(Class<T> clazz) throws Exception {
		// 取属性上的自定义特性
		for (Field field : fields) {
			if (field.getAnnotations() != null && field.getAnnotations().length > 0) {
				// 获取所有标有注解的列
				if (field.isAnnotationPresent(Display.class)) {
					Display dn = (Display) field.getAnnotation(Display.class);
					String name = dn.value();
					mapDisplayName.put(field, name);
				}
				if (field.isAnnotationPresent(Align.class)) {
					Align eca = (Align) field.getAnnotation(Align.class);
					EnumCellAlign align = eca.value();
					switch (align) {
					case LEFT:
						mapAlign.put(field, jxl.format.Alignment.LEFT);
						break;
					case CENTER:
						mapAlign.put(field, jxl.format.Alignment.CENTRE);
						break;
					case RIGHT:
						mapAlign.put(field, jxl.format.Alignment.RIGHT);
						break;
					}
				}
				if (field.isAnnotationPresent(Index.class)) {
					Index eci = field.getAnnotation(Index.class);
					int index = eci.value();
					mapColumnIndex.put(field, index);
				}
				if (field.isAnnotationPresent(Width.class)) {
					Width ecw = field.getAnnotation(Width.class);
					int width = ecw.value();
					mapWidth.put(field, width);
				}
				if (field.isAnnotationPresent(Converter.class)) {
					Converter ec = field.getAnnotation(Converter.class);
					Class<? extends IValueConverter> cc = ec.value();
					IValueConverter converter = cc.newInstance();
					mapValueConverter.put(field, converter);
				}
				if (field.isAnnotationPresent(Color.class)) {
					Color ec = field.getAnnotation(Color.class);
					Class<? extends IColorPicker> cc = ec.value();
					IColorPicker converter = cc.newInstance();
					mapColorConverter.put(field, converter);
				}
				if (field.isAnnotationPresent(Indexs.class)) {
					Indexs ci = field.getAnnotation(Indexs.class);
					int[] indexs = ci.value();
					loopTime = indexs.length;
					Integer[] ints = new Integer[loopTime];
					for (int i = 0; i < loopTime; i++) {
						ints[i] = new Integer(indexs[i]);
					}
					mapColumnIndexs.put(field, ints);
				}
			}
		}
	}

	/**
	 * 将颜色字符串转化为Colour
	 * 
	 * @param colorString
	 *            例如:#123456
	 * @return
	 */
	private Colour getNearestColour(String colorString) {
		return getNearestColour(java.awt.Color.decode(colorString));
	}

	/**
	 * 转换颜色对象
	 * 
	 * @param awtColor
	 * @return
	 */
	private Colour getNearestColour(java.awt.Color awtColor) {
		Colour color = null;

		Colour[] colors = Colour.getAllColours();
		if ((colors != null) && (colors.length > 0)) {
			Colour crtColor = null;
			int[] rgb = null;
			int diff = 0;
			int minDiff = 999;

			for (int i = 0; i < colors.length; i++) {
				crtColor = colors[i];
				rgb = new int[3];
				rgb[0] = crtColor.getDefaultRGB().getRed();
				rgb[1] = crtColor.getDefaultRGB().getGreen();
				rgb[2] = crtColor.getDefaultRGB().getBlue();

				diff = Math.abs(rgb[0] - awtColor.getRed()) + Math.abs(rgb[1] - awtColor.getGreen())
						+ Math.abs(rgb[2] - awtColor.getBlue());

				if (diff < minDiff) {
					minDiff = diff;
					color = crtColor;
				}
			}
		}
		if (color == null)
			color = Colour.WHITE;
		return color;
	}
}
