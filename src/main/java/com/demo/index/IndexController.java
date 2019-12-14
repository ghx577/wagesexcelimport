package com.demo.index;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com.demo.common.model.Wages;
import com.jfinal.core.Controller;
import cn.hutool.core.io.FileUtil;
import cn.hutool.core.util.IdUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;

/**
 * 本 demo 仅表达最为粗浅的 jfinal 用法，更为有价值的实用的企业级用法 详见 JFinal 俱乐部:
 * http://jfinal.com/club
 * 
 * IndexController
 */
public class IndexController extends Controller {
	public void index() {

		ExcelReader reader = ExcelUtil.getReader(FileUtil.file("D:\\demo\\demo.xls"));

		List<Sheet> sheets = reader.getSheets();
		System.out.println("sheet的表格数是：" + sheets.size());
		int num = 1;
		for (Sheet sheet : sheets) {
			int i = num++;
			System.out.println("执行第几个表格：" + i);
			// 遍历行Row
			// 先要定位到姓名 卡号 金额 银行 的列数 ，获取到列数后然后把数据逐个抓取
			int xuhaonum = 0;
			int xmnum = 1;
			int jenum = 0;
			int khnum = 0;
			int yhnum = 0;

			// 遍历获取
			int celllastnum = 0;
			for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
				Row hssfRow = sheet.getRow(rowNum);
				if (hssfRow == null) {
					continue;
				}
				// 求最大的列数
				//System.out.println("第" + i + "表第"+rowNum+"行的最大列数:" + hssfRow.getLastCellNum());
				//最大的必须有值
				int lastnum = hssfRow.getLastCellNum();
				if(StrUtil.isNotEmpty(getValue(hssfRow.getCell(lastnum-1)))) {
					celllastnum = hssfRow.getLastCellNum() > celllastnum ? hssfRow.getLastCellNum() : celllastnum;
				}
			}
			jenum = celllastnum -4;
			khnum = celllastnum -2;
			yhnum = celllastnum-1;
			System.out.println("第" + i + ":" + xmnum);
			System.out.println("第" + i + ":" + jenum);
			System.out.println("第" + i + ":" + khnum);
			System.out.println("第" + i + ":" + yhnum);

			// 循环行列获取数据
			for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
				Row hssfRow = sheet.getRow(rowNum);
				if (hssfRow == null) {
					continue;
				}

				String xuhao = getValue(hssfRow.getCell(xuhaonum));
				if(StrUtil.isNotEmpty(xuhao)&&!"序号".equals(xuhao)&&xuhao.length()<=5) {
					String xm = getValue(hssfRow.getCell(xmnum));
					String je = getValue(hssfRow.getCell(jenum));
					String kh = getValue(hssfRow.getCell(khnum));
					String yh = getValue(hssfRow.getCell(yhnum));
					
					Wages pojo = new Wages();
					pojo.setId(IdUtil.fastSimpleUUID());
					pojo.setCreateTime(new Date());
					pojo.setUsername(xm);
					pojo.setAmount(new BigDecimal(je));
					pojo.setCarno(kh);
					pojo.setBank(yh);
					pojo.save();
					
					System.out.println("第" + i + ":" + xm+":"+ je+":"+ kh+":"+ yh);
				}
			}

		}

		render("index.html");
	}

	private static String getValue(Cell hssfCell) {
		String value = "";

		if (hssfCell != null) {
			if (hssfCell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
				value = String.valueOf(hssfCell.getBooleanCellValue());
			} else if (hssfCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
				value = String.valueOf(hssfCell.getNumericCellValue());
			} else if (hssfCell.getCellType() == Cell.CELL_TYPE_FORMULA) {
				value = String.valueOf(hssfCell.getNumericCellValue());
			} else {
				value = String.valueOf(hssfCell.getStringCellValue());
			}
		}
		value = StrUtil.isNotEmpty(value) ? value.trim() : value;
		return value;
	}

}
