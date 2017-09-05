package com.bjhy.inline.office.excel;

import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 将excel转成html
 * 
 * @author wubo
 *
 */
public class ExcelToHtml {

	/**
	 * 程序入口方法
	 * 
	 * @param fileStream 文件流
	 * @param isWithStyle 是否需要表格样式 包含 字体 颜色 边框 对齐方式
	 * @return <table>
	 *         ...
	 *         </table>
	 *         字符串
	 * @throws InlineOfficeException 
	 */
	public String readExcelToHtml(InputStream fileStream, boolean isWithStyle){

		String htmlExcel = null;
		try {
			Workbook wb = WorkbookFactory.create(fileStream);
			if (wb instanceof XSSFWorkbook) {
				XSSFWorkbook xWb = (XSSFWorkbook) wb;
				htmlExcel = getExcelInfo(xWb, isWithStyle);
			} else if (wb instanceof HSSFWorkbook) {
				HSSFWorkbook hWb = (HSSFWorkbook) wb;
				htmlExcel = getExcelInfo(hWb, isWithStyle);
			}
		} catch (Exception e) {
			e.printStackTrace();
			throw new RuntimeException("当前文件不是标准的excel文件,无法打开");
		} finally {
			try {
				fileStream.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return htmlExcel;
	}

	public String getExcelInfo(Workbook wb, boolean isWithStyle) {
		StringBuffer sb = new StringBuffer();

		int sheetNum = wb.getNumberOfSheets();
//		 int sheetNum = 6;//发现问题，调用用
		// 遍历每一个sheet
		for (int sheetIndex = 0; sheetIndex < sheetNum; sheetIndex++) {
			Sheet sheet = wb.getSheetAt(sheetIndex);// 获取第一个Sheet的内容
			int lastRowNum = sheet.getLastRowNum();
			Map<String, String> map[] = getRowSpanColSpanMap(sheet);
			sb.append("<table  style='border-collapse:collapse;' width='100%'>");
			Row row = null; // 兼容
			Cell cell = null; // 兼容
			int maxcell = 0;
			for (int i = 1; i < lastRowNum; i++) {
				row = sheet.getRow(i);
				if (row != null) {
					if (maxcell < row.getLastCellNum()) {
						maxcell = row.getLastCellNum();
					}
				}

			}
			for (int rowNum = sheet.getFirstRowNum(); rowNum <= lastRowNum; rowNum++) {
				row = sheet.getRow(rowNum);
				sb.append("<tr >");
				if (row == null) {
					for (int colNum = 0; colNum < maxcell; colNum++) {
						
						if (map[0].containsKey(rowNum + "," + colNum)) {
							String pointString = map[0].get(rowNum + "," + colNum);
							map[0].remove(rowNum + "," + colNum);
							int bottomeRow = Integer.valueOf(pointString.split(",")[0]);
							int bottomeCol = Integer.valueOf(pointString.split(",")[1]);
							int rowSpan = bottomeRow;
							int colSpan = bottomeCol;
							sb.append("<td rowspan= '" + rowSpan + "' colspan= '"+ colSpan + "' ");
						} else if (map[1].containsKey(rowNum + "," + colNum)) {
							map[1].remove(rowNum + "," + colNum);
							continue;
						} else {
							sb.append("<td ");
						}
						
						sb.append(" style='background-color: white;border: 1px solid #D0D7E5;'> &nbsp;</td>");
					}
					sb.append("</tr>");
					continue;
				}
				for (int colNum = 0; colNum < maxcell; colNum++) {
					cell = row.getCell(colNum);
					
					if (map[0].containsKey(rowNum + "," + colNum)) {
						String pointString = map[0].get(rowNum + "," + colNum);
						map[0].remove(rowNum + "," + colNum);
						int bottomeRow = Integer.valueOf(pointString.split(",")[0]);
						int bottomeCol = Integer.valueOf(pointString.split(",")[1]);
						int rowSpan = bottomeRow;
						int colSpan = bottomeCol;
						sb.append("<td rowspan= '" + rowSpan + "' colspan= '"+ colSpan + "' ");
					} else if (map[1].containsKey(rowNum + "," + colNum)) {
						map[1].remove(rowNum + "," + colNum);
						continue;
					} else {
						sb.append("<td ");
					}

					// 判断是否需要样式
					
					if (cell == null) { // 特殊情况 空白的单元格会返回null
						// 判断是否需要样式  //这里没有让样式一定就一致,只是尽量让其一致
						if (isWithStyle) { 
							cell = getNotEmptyStyleCell(sheet,row, rowNum,cell, colNum,colNum, maxcell);
							if(cell == null){
								sb.append(" style='background-color: white;border: 1px solid #D0D7E5;'>  &nbsp;</td>");
							}else{
								sb.append(" ");
								dealExcelStyle(wb, sheet, cell, sb);// 处理单元格样式
								sb.append("> &nbsp;</td>");
								System.out.println();
							}
							
						}else{
							sb.append("<td style='background-color: white;border: 1px solid #D0D7E5;'>  &nbsp;</td>");
						}
						continue;
					}else{
						String stringValue = getCellValue(cell);
						//自动换行
						if(stringValue != null){
							stringValue = stringValue.replace("\n", "<br>");
						}
						
						if (isWithStyle) {
							dealExcelStyle(wb, sheet, cell, sb);// 处理单元格样式
						}

						sb.append(">");
						if (stringValue == null || "".equals(stringValue.trim())) {
							sb.append(" &nbsp; ");
						} else {
							sb.append(stringValue.replace(String.valueOf((char) 160), "&nbsp;"));
						}
						sb.append("</td>");
					}
				}
				sb.append("</tr>");
			}
			sb.append("</table>");
		}

		return sb.toString();
	}
	
	/**
	 * 当 initColNum 不为零时,从当前列向上找,一直找到不为空的节点,若最终还是为空那就返回空cell,若 initColNum 为零则相反找
	 * @param sheet 当前sheet
	 * @param row 当前行
	 * @param rowNum 当前行的index
	 * @param cell 当前 null 列
	 * @param initColNum 初始列的index
	 * @param colNum 当前列的index
	 * @param lastColNum 最大的index
	 * @return
	 */
	private Cell getNotEmptyStyleCell(Sheet sheet,Row row,int rowNum,Cell cell,int initColNum,int colNum,int lastColNum){
		if(initColNum > 0 && row.getCell(0) != null){
			cell = row.getCell(colNum-1);
			if(cell == null){
				colNum--;
				cell = getNotEmptyStyleCell(sheet,row, rowNum,cell, initColNum,colNum, lastColNum);
			}
		}else if(colNum<=lastColNum){
			cell = row.getCell(colNum+1);
			if(cell == null){
				colNum++;
				cell = getNotEmptyStyleCell(sheet,row, rowNum,cell, initColNum,colNum, lastColNum);
			}
		}
		return cell;
	}

	@SuppressWarnings({ "unchecked", "rawtypes" })
	private Map<String, String>[] getRowSpanColSpanMap(Sheet sheet) {

		Map<String, String> map0 = new HashMap<String, String>();
		Map<String, String> map1 = new HashMap<String, String>();
		int mergedNum = sheet.getNumMergedRegions();
		CellRangeAddress range = null;
		for (int i = 0; i < mergedNum; i++) {
			range = sheet.getMergedRegion(i);
			int topRow = range.getFirstRow();
			int topCol = range.getFirstColumn();
			int bottomRow = range.getLastRow();
			int bottomCol = range.getLastColumn();
//			map0.put(topRow + "," + topCol, bottomRow + "," + bottomCol);
			map0.put(topRow + "," + topCol, (bottomRow- topRow + 1)+ "," + (bottomCol - topCol + 1));
			int tempRow = topRow;
			while (tempRow <= bottomRow) {
				int tempCol = topCol;
				while (tempCol <= bottomCol) {
					map1.put(tempRow + "," + tempCol, "");
					tempCol++;
				}
				tempRow++;
			}
			map1.remove(topRow + "," + topCol);
		}
		Map[] map = { map0, map1 };
		return map;
	}
	
	// 获得合并单元格的区域
	public Map<String, String> getMergeMap(XSSFSheet currentSheet) {
		Map<String, String> mergeMap = new HashMap<String, String>();
		int num = currentSheet.getNumMergedRegions();
		for (int index = 0; index < num; index++) {
			CellRangeAddress cra = currentSheet.getMergedRegion(index);
			int firstColumn = cra.getFirstColumn();
			int lastColumn = cra.getLastColumn();
			int firstRow = cra.getFirstRow();
			int lastRow = cra.getLastRow();
			mergeMap.put(firstRow + "," + firstColumn, (lastColumn- firstColumn + 1)+ "," + (lastRow - firstRow + 1));
		}
		return mergeMap;
	}

	/**
	 * 获取表格单元格Cell内容
	 * 
	 * @param cell
	 * @return
	 */
	private String getCellValue(Cell cell) {

		String result = new String();
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_NUMERIC:// 数字类型
			if (HSSFDateUtil.isCellDateFormatted(cell)) {// 处理日期格式、时间格式
				SimpleDateFormat sdf = null;
				if (cell.getCellStyle().getDataFormat() == HSSFDataFormat
						.getBuiltinFormat("h:mm")) {
					sdf = new SimpleDateFormat("HH:mm");
				} else {// 日期
					sdf = new SimpleDateFormat("yyyy-MM-dd");
				}
				Date date = cell.getDateCellValue();
				result = sdf.format(date);
			} else if (cell.getCellStyle().getDataFormat() == 58) {
				// 处理自定义日期格式：m月d日(通过判断单元格的格式id解决，id的值是58)
				SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
				double value = cell.getNumericCellValue();
				Date date = org.apache.poi.ss.usermodel.DateUtil
						.getJavaDate(value);
				result = sdf.format(date);
			} else {
				double value = cell.getNumericCellValue();
				CellStyle style = cell.getCellStyle();
				DecimalFormat format = new DecimalFormat();
				String temp = style.getDataFormatString();
				// 单元格设置成常规
				if ("General".equals(temp)) {
					format.applyPattern("#");
				}
				result = format.format(value);
			}
			break;
		case Cell.CELL_TYPE_STRING:// String类型
			result = cell.getRichStringCellValue().toString();
			break;
		case Cell.CELL_TYPE_FORMULA:
//			result = cell.getCellFormula();
			result ="";
			break;
		case Cell.CELL_TYPE_BLANK:
			result = "";
			break;
		default:
			result = "";
			break;
		}
		return result;
	}

	/**
	 * 处理表格样式
	 * 
	 * @param wb
	 * @param sheet
	 * @param cell
	 * @param sb
	 */
	private void dealExcelStyle(Workbook wb, Sheet sheet, Cell cell,
			StringBuffer sb) {
		CellStyle cellStyle = cell.getCellStyle();
		if (cellStyle != null) {
			short alignment = cellStyle.getAlignment();
			sb.append("align='" + convertAlignToHtml(alignment) + "' ");// 单元格内容的水平对齐方式
			short verticalAlignment = cellStyle.getVerticalAlignment();
			sb.append("valign='"
					+ convertVerticalAlignToHtml(verticalAlignment) + "' ");// 单元格中内容的垂直排列方式
			if (wb instanceof XSSFWorkbook) {
				XSSFFont xf = ((XSSFCellStyle) cellStyle).getFont();
				short boldWeight = xf.getBoldweight();
				sb.append("style='");
				sb.append("font-weight:" + boldWeight + ";"); // 字体加粗
				sb.append("font-size: " + xf.getFontHeight() / 2 + "%;"); // 字体大小
				int columnWidth = sheet.getColumnWidth(cell.getColumnIndex());
				sb.append("width:" + columnWidth + "px;"); 
				XSSFColor xc = xf.getXSSFColor();
				if (xc != null && !"".equals(xc)) {
					sb.append("color:#" + xc.getARGBHex().substring(2) + ";"); // 字体颜色
				}
				XSSFColor bgColor = (XSSFColor) cellStyle
						.getFillForegroundColorColor();
				if (bgColor != null && !"".equals(bgColor)) {
					sb.append("background-color:#"+ bgColor.getARGBHex().substring(2) + ";"); // 背景颜色
				}//0070C0
				sb.append(getBorderStyle(0, cellStyle.getBorderTop(),((XSSFCellStyle) cellStyle).getTopBorderXSSFColor()));
				sb.append(getBorderStyle(1, cellStyle.getBorderRight(),((XSSFCellStyle) cellStyle).getRightBorderXSSFColor()));
				sb.append(getBorderStyle(2, cellStyle.getBorderBottom(),((XSSFCellStyle) cellStyle).getBottomBorderXSSFColor()));
				sb.append(getBorderStyle(3, cellStyle.getBorderLeft(),((XSSFCellStyle) cellStyle).getLeftBorderXSSFColor()));
			} else if (wb instanceof HSSFWorkbook) {
				HSSFFont hf = ((HSSFCellStyle) cellStyle).getFont(wb);
				short boldWeight = hf.getBoldweight();
				short fontColor = hf.getColor();
				sb.append("style='");
				HSSFPalette palette = ((HSSFWorkbook) wb).getCustomPalette(); // 类HSSFPalette用于求的颜色的国际标准形式
				HSSFColor hc = palette.getColor(fontColor);
				sb.append("font-weight:" + boldWeight + ";"); // 字体加粗
				sb.append("font-size: " + hf.getFontHeight() / 2 + "%;"); // 字体大小
				String fontColorStr = convertToStardColor(hc);
				if (fontColorStr != null && !"".equals(fontColorStr.trim())) {
					sb.append("color:" + fontColorStr + ";"); // 字体颜色
				}
				int columnWidth = sheet.getColumnWidth(cell.getColumnIndex());
				sb.append("width:" + columnWidth + "px;");
				short bgColor = cellStyle.getFillForegroundColor();
				hc = palette.getColor(bgColor);
				String bgColorStr = convertToStardColor(hc);
				if (bgColorStr != null && !"".equals(bgColorStr.trim())) {
					sb.append("background-color:" + bgColorStr + ";"); // 背景颜色
				}
				sb.append(getBorderStyle(palette, 0, cellStyle.getBorderTop(),cellStyle.getTopBorderColor()));
				sb.append(getBorderStyle(palette, 1,cellStyle.getBorderRight(),cellStyle.getRightBorderColor()));
				sb.append(getBorderStyle(palette, 3, cellStyle.getBorderLeft(),cellStyle.getLeftBorderColor()));
				sb.append(getBorderStyle(palette, 2,cellStyle.getBorderBottom(),cellStyle.getBottomBorderColor()));
			}
			sb.append("' ");
		} else {

		}

	}

	/**
	 * 单元格内容的水平对齐方式
	 * 
	 * @param alignment
	 * @return
	 */
	private String convertAlignToHtml(short alignment) {

		String align = "left";
		switch (alignment) {
		case CellStyle.ALIGN_LEFT:
			align = "left";
			break;
		case CellStyle.ALIGN_CENTER:
			align = "center";
			break;
		case CellStyle.ALIGN_RIGHT:
			align = "right";
			break;
		default:
			break;
		}
		return align;
	}

	/**
	 * 单元格中内容的垂直排列方式
	 * 
	 * @param verticalAlignment
	 * @return
	 */
	private String convertVerticalAlignToHtml(short verticalAlignment) {

		String valign = "middle";
		switch (verticalAlignment) {
		case CellStyle.VERTICAL_BOTTOM:
			valign = "bottom";
			break;
		case CellStyle.VERTICAL_CENTER:
			valign = "center";
			break;
		case CellStyle.VERTICAL_TOP:
			valign = "top";
			break;
		default:
			break;
		}
		return valign;
	}

	private String convertToStardColor(HSSFColor hc) {

		StringBuffer sb = new StringBuffer("");
		if (hc != null) {
			if (HSSFColor.AUTOMATIC.index == hc.getIndex()) {
				return null;
			}
			sb.append("#");
			for (int i = 0; i < hc.getTriplet().length; i++) {
				sb.append(fillWithZero(Integer.toHexString(hc.getTriplet()[i])));
			}
		}

		return sb.toString();
	}

	private String fillWithZero(String str) {
		if (str != null && str.length() < 2) {
			return "0" + str;
		}
		return str;
	}

	String[] bordesr = { "border-top:", "border-right:", "border-bottom:",
			"border-left:" };
	String[] borderStyles = { "solid ", "solid ", "solid ", "solid ", "solid ",
			"solid ", "solid ", "solid ", "solid ", "solid", "solid", "solid",
			"solid", "solid" };

	private String getBorderStyle(HSSFPalette palette, int b, short s, short t) {

		if (s == 0)
			return bordesr[b] + borderStyles[s] + "#d0d7e5 1px;";
		;
		String borderColorStr = convertToStardColor(palette.getColor(t));
		borderColorStr = borderColorStr == null || borderColorStr.length() < 1 ? "#000000"
				: borderColorStr;
		return bordesr[b] + borderStyles[s] + borderColorStr + " 1px;";

	}

	private String getBorderStyle(int b, short s, XSSFColor xc) {

		
		if (xc != null && !"".equals(xc)) {
			String borderColorStr = xc.getARGBHex();// t.getARGBHex();
			borderColorStr = borderColorStr == null|| borderColorStr.length() < 1 ? "#000000" : "#"+borderColorStr.substring(2);
			return bordesr[b] + borderStyles[s] + borderColorStr + " 1px;";
		}
		
		if (s == 0)
			return bordesr[b] + borderStyles[s] + "#d0d7e5 1px;";

		return "";
	}
	
}
