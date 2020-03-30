package com.github.util;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import cn.afterturn.easypoi.word.WordExportUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import com.github.entity.FilePath;
import com.github.entity.TableMap;
import com.github.entity.WordTable;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import javax.validation.constraints.NotNull;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigInteger;
import java.net.URLEncoder;
import java.util.List;
import java.util.Map;


/**
 *
 * 利用hutool和easypoi工具包装的自定义excel操作工具
 *
 * @author 王俊锋
 * @version  2020/03/05 v1.0
 */
public class MyOfficeUtil {

	/**
	 *
	 * @param path 文件所在路径
	 * @param headerAlias K-V值对应reader.addHeaderAlias()方法的两个参数
	 * @param clazz 指定类型的class
	 * @param <T> 指定类型的clazz
	 * @return 生成指定类型的list
	 */
	public static <T> List<T> getList(@NotNull String path, @NotNull Map<String, String> headerAlias, @NotNull Class<T> clazz) {
		ExcelReader reader = ExcelUtil.getReader(path);
		headerAlias.forEach(reader::addHeaderAlias);
		return reader.readAll(clazz);
	}

	/**
	 * @param templateFilePath 模板文件路径
	 * @param newFilePath 新生成的文件保存路径
	 * @param newFileName 新生成文件名称
	 * @param map           生成文件数据
	 * @return 返回了路径对象（包含了真实路径和相对路径）
	 */
	public static FilePath getWordFile(@NotNull String templateFilePath, @NotNull String newFilePath, @NotNull String newFileName, @NotNull Map<String, Object> map) {
		File file = new File(templateFilePath);
		if (!file.exists()) {
			throw new NullPointerException("模板文件:" + templateFilePath + "不存在！");
		}
		String path;
		try {
			path = newFilePath + "/system-create/word/";
			File fileDir = new File(path);

			if (!fileDir.exists()) {
				fileDir.mkdirs();
			}

			XWPFDocument doc = WordExportUtil.exportWord07(templateFilePath, map);
			FileOutputStream fos = new FileOutputStream(path + newFileName);
			doc.write(fos);
			fos.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return new FilePath(newFilePath + "/system-create/word/" + newFileName, "/system-create/word/" + newFileName);
	}

	/**
	 * @param templateFilePath 模板文件路径
	 * @param newFilePath 新生成的文件路径
	 * @param newFileName 新生成的文件名
	 * @param map 生成的文件数据
	 * @param response 将生成的文件直接写入HttpServletResponse中
	 */
	public static void getWordFile(@NotNull String templateFilePath, @NotNull String newFilePath, @NotNull String newFileName, @NotNull Map<String, Object> map, HttpServletResponse response) {
		File file = new File(templateFilePath);
		if (!file.exists()) {
			throw new NullPointerException("模板文件:" + templateFilePath + "不存在！");
		}

		try {
			XWPFDocument doc = WordExportUtil.exportWord07(templateFilePath, map);
			String path = newFilePath + "/system-create/word/";
			FileOutputStream fos = new FileOutputStream(path + newFileName);
			doc.write(fos);
			fos.close();

			ServletOutputStream outputStream = response.getOutputStream();
			// 设置强制下载不打开
			response.setContentType("application/force-download");
			response.setHeader("content-disposition", "attachment;filename=" + URLEncoder.encode(newFileName, "UTF-8"));
			doc.write(outputStream);
			outputStream.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * 生成excel文件并返回其路径
	 * @param templateFilePath 模版文件路径
	 * @param newFilePath 生成的新文件路径
	 * @param newFileName 生成的新文件名
	 * @param map 生成的文件信息
	 * @return 返回了文件路径对象
	 * @throws IOException 文件及IO流异常
	 */
	public static FilePath getExcelFile(@NotNull String templateFilePath, @NotNull String newFilePath, @NotNull String newFileName, @NotNull Map<String, Object> map) throws IOException {
		File templateFile = new File(templateFilePath);
		if (!templateFile.exists()) {
			throw new NullPointerException("模板文件不存在！");
		}
		TemplateExportParams params = new TemplateExportParams(templateFilePath);
		Workbook workbook = ExcelExportUtil.exportExcel(params, map);
		File newPath = new File(newFilePath + "/system-create/excel/");
		if (!newPath.exists()) {
			newPath.mkdirs();
		}
		String file = newFilePath + "/system-create/excel/" + newFileName;
		FileOutputStream out = new FileOutputStream(file);
		workbook.write(out);
		out.close();

		return new FilePath(newFilePath + "/system-create/excel/" + newFileName, "/system-create/excel/" + newFileName);
	}

	/**
	 * 生成文件并写入输出流中
	 *
	 * @param templateFilePath 模版文件路径
	 * @param newFilePath      生成的新文件路径
	 * @param newFileName      新的文件名
	 * @param map              生成文件的数据
	 * @param response         写进的流
	 * @throws IOException 文件及IO流异常
	 */
	public static void getExcelFile(@NotNull String templateFilePath, @NotNull String newFilePath, @NotNull String newFileName, @NotNull Map<String, Object> map, @NotNull HttpServletResponse response) throws IOException {
		File templateFile = new File(templateFilePath);
		if (!templateFile.exists()) {
			throw new NullPointerException("模板文件不存在！");
		}

		File newPath = new File(newFilePath + "/system-create/excel/");
		if (!newPath.exists()) {
			newPath.mkdirs();
		}

		TemplateExportParams params = new TemplateExportParams(templateFilePath);
		Workbook workbook = ExcelExportUtil.exportExcel(params, map);
//		在服务器本地生成文件
		String file = newFilePath + "/system-create/excel/" + newFileName;
		FileOutputStream fos = new FileOutputStream(file);
		workbook.write(fos);
		fos.close();

//		直接将文件写进流中
		ServletOutputStream out = response.getOutputStream();
		// 设置强制下载不打开
		response.setContentType("application/force-download");
		response.setHeader("content-disposition", "attachment;filename=" + URLEncoder.encode(newFileName, "UTF-8"));
		workbook.write(out);
		out.close();
	}

	/**
	 * 生成word表格并合并列或者行
	 *
	 * @param newFilePath       生成的新文件路径
	 * @param newFileName       生成的新文件名
	 * @param count             合并第几列/行
	 * @param from              从第几行/列开始
	 * @param to                到第几行/列结束
	 * @param map               表头与实体的映射
	 * @param list              实体信息
	 * @param isReverse         是否倒序
	 * @param clazz             实体的class
	 * @param <T>               实体类
	 * @param isEnjambmentMerge 是否跨行合并单元格,true为跨行合并,false为跨列合并
	 * @return 返回生成的文件路径（相对于newFilePath参数，newFilePath+返回路径=真实路径）
	 * @throws IOException               文件及IO流异常
	 * @throws NoSuchFieldException      未找到实体类字段
	 * @throws NoSuchMethodException     未找到实体类方法
	 * @throws InvocationTargetException 反射执行方法异常（调用属性get方法失败）
	 * @throws IllegalAccessException    反射执行方法异常（调用属性get方法失败）
	 */
	public static <T> FilePath getWordFile(String newFilePath, String newFileName, int count, int from, int to,
	                                       Map<Integer, TableMap> map, List<? extends T> list, boolean isReverse,
	                                       Class<T> clazz, boolean isEnjambmentMerge)
			throws IOException, NoSuchFieldException, NoSuchMethodException,
			InvocationTargetException, IllegalAccessException {

		String file = newFilePath + "/system-create/excel/" + newFileName;

		int rowSize = list.size();

		WordTable wordInfo = getWordInfo(rowSize, isReverse, map, list, clazz);

		if (isEnjambmentMerge) {
			mergeCellsVertically(wordInfo.getTable(), count, from, to);
		} else {
			mergeCellsHorizontal(wordInfo.getTable(), count, from, to);
		}

		FileOutputStream fos = new FileOutputStream(file);
		wordInfo.getDocument().write(fos);
		fos.close();

		return new FilePath(file, "/system-create/excel/" + newFileName);
	}

	/**
	 * 获取word表格
	 *
	 * @param newFilePath 生成的新文件路径
	 * @param newFileName 生成的新文件名
	 * @param map         表头与实体的映射
	 * @param list        实体信息
	 * @param isReverse   是否倒序
	 * @param clazz       实体的class
	 * @param <T>         实体类
	 * @return 返回生成的文件路径（相对于newFilePath参数，newFilePath+返回路径=真实路径）
	 * @throws IOException               文件及IO流异常
	 * @throws NoSuchFieldException      未找到实体类字段
	 * @throws NoSuchMethodException     未找到实体类方法
	 * @throws InvocationTargetException 反射执行方法异常（调用属性get方法失败）
	 * @throws IllegalAccessException    反射执行方法异常（调用属性get方法失败）
	 */
	public static <T> FilePath getWordFile(String newFilePath, String newFileName,
	                                       Map<Integer, TableMap> map, List<? extends T> list, boolean isReverse,
	                                       Class<T> clazz)
			throws IOException, NoSuchFieldException, NoSuchMethodException,
			InvocationTargetException, IllegalAccessException {

		String file = newFilePath + "/system-create/excel/" + newFileName;

		int rowSize = list.size();

		WordTable wordInfo = getWordInfo(rowSize, isReverse, map, list, clazz);

		FileOutputStream fos = new FileOutputStream(file);
		wordInfo.getDocument().write(fos);
		fos.close();

		return new FilePath(file, "/system-create/excel/" + newFileName);
	}

	/**
	 * 跨行合并单元格
	 *
	 * @param table   合并的单元格的表格
	 * @param col     合并第几列
	 * @param fromRow 从第几行开始合并
	 * @param toRow   合并到第几列
	 */
	private static void mergeCellsVertically(XWPFTable table, int col, int fromRow, int toRow) {
		for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
			XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
			if (rowIndex == fromRow) {
				cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
			} else {
				cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
			}
		}
	}

	/**
	 * 跨列合并单元格
	 *
	 * @param table    合并的单元格的表格
	 * @param row      合并第几行
	 * @param fromCell 从第几个单元格开始
	 * @param toCell   到第几个单元格
	 */
	private static void mergeCellsHorizontal(XWPFTable table, int row, int fromCell, int toCell) {
		for (int cellIndex = fromCell; cellIndex <= toCell; cellIndex++) {
			XWPFTableCell cell = table.getRow(row).getCell(cellIndex);
			if (cellIndex == fromCell) {
				cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
			} else {
				cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
			}
		}
	}

	/**
	 * 生成并返回word表格信息
	 *
	 * @param rowSize   生成的表格一共有多少列，除表头外
	 * @param isReverse 是否倒序
	 * @param map       表头与实体字段的映射关系
	 * @param list      实体信息
	 * @param clazz     所处理的实体的class
	 * @param <T>       实体类型
	 * @return 返回新建的word对象
	 * @throws NoSuchFieldException      未找到实体类字段
	 * @throws NoSuchMethodException     未找到实体类方法
	 * @throws InvocationTargetException 反射执行方法异常（调用属性get方法失败）
	 * @throws IllegalAccessException    反射执行方法异常（调用属性get方法失败）
	 */
	private static <T> WordTable getWordInfo(int rowSize, boolean isReverse, Map<Integer, TableMap> map, List<? extends T> list, Class<T> clazz) throws NoSuchFieldException, NoSuchMethodException, InvocationTargetException, IllegalAccessException {
//		新建一个word对象
		XWPFDocument doc = new XWPFDocument();

		CTDocument1 document = doc.getDocument();
		CTBody body = document.getBody();
		if (!body.isSetSectPr()) {
			body.addNewSectPr();
		}
		CTSectPr section = body.getSectPr();

		if (!section.isSetPgSz()) {
			section.addNewPgSz();
		}
		CTPageSz pageSize = section.getPgSz();
		pageSize.setW(BigInteger.valueOf(15840));
		pageSize.setH(BigInteger.valueOf(12240));
		pageSize.setOrient(STPageOrientation.LANDSCAPE);

		//添加标题
		doc.createParagraph();

		//表格
		XWPFTable ComTable = doc.createTable();

		//设置指定宽度
		CTTbl ttbl = ComTable.getCTTbl();
		CTTblGrid tblGrid = ttbl.addNewTblGrid();
		CTTblGridCol gridCol = tblGrid.addNewGridCol();
		gridCol.setW(new BigInteger(800 + ""));

		//表头
		XWPFTableRow rowHead = ComTable.getRow(0);

		int k = 0;

		for (Map.Entry<Integer, TableMap> ignored : map.entrySet()) {
			if (rowHead.getCell(k) == null) {
				rowHead.addNewTableCell();
			}
			k++;
		}

		int k1 = 0;
		if (isReverse) {
			k1 = map.entrySet().size() - 1;
		}

		for (int i = 0; i < map.entrySet().size(); i++) {
			XWPFRun xwpfRun = getXWPFRun(isReverse, i, k1, rowHead);
			xwpfRun.setBold(true);
			xwpfRun.setText(map.get(i).getKey());
		}

		//表格内容
		for (T t : list) {
			XWPFTableRow rowsContent = ComTable.createRow();
			for (int j = 0; j < map.entrySet().size(); j++) {
				XWPFRun xwpfRun = getXWPFRun(isReverse, j, k1, rowsContent);

				String fieldName = clazz.getDeclaredField(map.get(j).getValue()) + "";

				String substring = fieldName.substring(fieldName.lastIndexOf(".") + 1);

				String upperChar = substring.substring(0, 1).toUpperCase();
				String anotherStr = substring.substring(1);
				String methodName = "get" + upperChar + anotherStr;
				Method method = clazz.getMethod(methodName);
				method.setAccessible(true);
				Object resultValue = method.invoke(t);
				xwpfRun.setText(resultValue + ""); //单元格段落加载内容
			}
		}
		if (rowSize == 0) {
			for (int i = 0; i < clazz.getFields().length; i++) {
				XWPFTableCell cell = ComTable.getRow(0).getCell(i);
				cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER); //垂直居中
			}
		} else {
			//设置居中
			for (int i = 0; i <= rowSize; i++) {
				for (int j = 0; j < clazz.getFields().length; j++) {
					XWPFTableCell cell = ComTable.getRow(i).getCell(j);
					cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER); //垂直居中
				}
			}
		}
		return new WordTable(ComTable, doc);
	}

	//	设置WORD表格的每一行的格式
	private static XWPFRun getXWPFRun(boolean isReverse, int i, int j, XWPFTableRow row) {
		XWPFParagraph cellParagraph = null;
		if (isReverse) {
			cellParagraph = row.getCell(j - i).getParagraphs().get(0);
		} else {
			cellParagraph = row.getCell(i).getParagraphs().get(0);
		}
		cellParagraph.setAlignment(ParagraphAlignment.CENTER); //设置表格居中
		XWPFRun cellParagraphRun = cellParagraph.createRun();
		cellParagraphRun.setFontSize(10); //设置表格居中
		return cellParagraphRun;
	}

}
