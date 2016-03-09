package com.nari.sung;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.util.Arrays;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 所需jar包 下载地址http://www.apache.org/dyn/closer.cgi/poi/release/bin/poi-bin-3.9-
 * 20121203.zip
 * 
 * poi-3.7-20101029.jar poi-ooxml-3.7-20101029.jar
 * poi-ooxml-schemas-3.7-20101029.jar xmlbeans-2.3.0.jar dom4j-1.6.1.jar
 * 
 * 
 */
public class ExcelUtil {
	private static Map<String, Integer> titles = null;
	private static DecimalFormat df = new DecimalFormat("0");
	private static DecimalFormat nf = new DecimalFormat("0.000");// 格式化数字

	public static void main(String[] args) {
		try {
			// transFile("E:/workspace/workspace_hyl/文件/原始表格/场东村/土地承包清册.xls",
			// "1,4");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@SuppressWarnings( { "finally", "deprecation" })
	public static int transFile(String fName, Map<String, String> inputMap)
			throws Exception {
		int flag = 1;

		try {
			// 1、一次性读取excel表格中的内容
			String sr = inputMap.get("SR");
			String[] array = sr.split(",");
			if (array == null || array.length != 2) {
				return 0;
			}
			File inFile = new File(fName);
			List<Map<String, Object>> list = readExcel(inFile, Integer
					.parseInt(array[0]), Integer.parseInt(array[1]));

			// 2、创建Excel文件转换结果目录和以当前时间命名的xlsx文件
			int cbfbm = 1;
			int cyxh = 1;
			int hzNum = 0;
			int dkljh = 1;

			File outFile = null;

			HSSFWorkbook wb = null;
			HSSFSheet fbfdcHssfSheet = null;
			HSSFSheet jtcydcHssfSheet = null;
			HSSFSheet cbfdcHssfSheet = null;
			HSSFSheet dkdcHssfSheet = null;
			int rowNum = 0;
			int fbfRowNum = 0;
			int cbfRowNum = 0;
			int dkRowNum = 0;
			HSSFRow row = null;
			HSSFCell cell = null;
			HSSFFont hssfFont = null;
			HSSFCellStyle cellStyle = null;

			String zm = null;
			String hz = null;
			String hzsfz = null;
			String tdsyzbh = null;

			File fdir = new File(File.separator);
			File f = new File(fdir, "Excel文件转换结果");
			if (!f.isDirectory()) {
				f.mkdir();
			}

			Map<String, Object> tempMap = list.get(0);
			// 获取组名
			zm = StringUtil.removeNull(tempMap.get("ZM"));
			if (zm.indexOf("组") != -1) {
				zm = zm.substring(0, zm.indexOf("组"));
			} else if (zm.indexOf(".") != -1) {
				zm = zm.substring(0, zm.indexOf("."));
			}

			zm = zeroPadder(zm, 2);

			for (int i = 0; i < list.size(); i++) {
				Map<String, Object> map = list.get(i);

				if (i == 0) {
					outFile = new File(f + File.separator + zm + "组.xlsx");

					wb = new HSSFWorkbook();

					// 设置表格颜色
					hssfFont = wb.createFont();
					hssfFont.setColor(HSSFFont.COLOR_RED);
					cellStyle = wb.createCellStyle();
					cellStyle.setFont(hssfFont);

					// 创建发包方调查sheet
					fbfdcHssfSheet = wb.createSheet("发包方调查表");

					row = fbfdcHssfSheet.createRow(0);
					cell = row.createCell((short) 0);
					cell.setCellValue(new HSSFRichTextString("发包方编码"));
					cell = row.createCell((short) 1);
					cell.setCellValue(new HSSFRichTextString("发包方名称"));
					cell = row.createCell((short) 2);
					cell.setCellValue(new HSSFRichTextString("发包方村组行政名称"));
					cell = row.createCell((short) 3);
					cell.setCellValue(new HSSFRichTextString("发包方负责人姓名"));
					cell = row.createCell((short) 4);
					cell.setCellValue(new HSSFRichTextString("负责人证件类型"));
					cell = row.createCell((short) 5);
					cell.setCellValue(new HSSFRichTextString("负责人证件号码"));
					cell = row.createCell((short) 6);
					cell.setCellValue(new HSSFRichTextString("联系电话"));
					cell = row.createCell((short) 7);
					cell.setCellValue(new HSSFRichTextString("发包方地址"));
					cell = row.createCell((short) 8);
					cell.setCellValue(new HSSFRichTextString("邮政编码"));
					cell = row.createCell((short) 9);
					cell.setCellValue(new HSSFRichTextString("调查员"));
					cell = row.createCell((short) 10);
					cell.setCellValue(new HSSFRichTextString("调查日期"));
					cell = row.createCell((short) 11);
					cell.setCellValue(new HSSFRichTextString("调查记事"));
					cell = row.createCell((short) 12);
					cell.setCellValue(new HSSFRichTextString("发包方原负责人姓名"));

					row = fbfdcHssfSheet.createRow(1);
					cell = row.createCell((short) 0);
					cell.setCellValue(new HSSFRichTextString("FBFBM"));
					cell = row.createCell((short) 1);
					cell.setCellValue(new HSSFRichTextString("FBFMC"));
					cell = row.createCell((short) 2);
					cell.setCellValue(new HSSFRichTextString("FBFXZMC"));
					cell = row.createCell((short) 3);
					cell.setCellValue(new HSSFRichTextString("FBFFZRXM"));
					cell = row.createCell((short) 4);
					cell.setCellValue(new HSSFRichTextString("FZRZJLX"));
					cell = row.createCell((short) 5);
					cell.setCellValue(new HSSFRichTextString("FZRZJHM"));
					cell = row.createCell((short) 6);
					cell.setCellValue(new HSSFRichTextString("LXDH"));
					cell = row.createCell((short) 7);
					cell.setCellValue(new HSSFRichTextString("FBFDZ"));
					cell = row.createCell((short) 8);
					cell.setCellValue(new HSSFRichTextString("YZBM"));
					cell = row.createCell((short) 9);
					cell.setCellValue(new HSSFRichTextString("FBFDCY"));
					cell = row.createCell((short) 10);
					cell.setCellValue(new HSSFRichTextString("FBFDCRQ"));
					cell = row.createCell((short) 11);
					cell.setCellValue(new HSSFRichTextString("FBFDCJS"));
					cell = row.createCell((short) 12);
					cell.setCellValue(new HSSFRichTextString("FBFYFZRXM"));

					fbfRowNum = 1;

					// 创建家庭成员调查sheet
					jtcydcHssfSheet = wb.createSheet("家庭成员调查表");

					row = jtcydcHssfSheet.createRow(0);
					cell = row.createCell((short) 0);
					cell.setCellValue(new HSSFRichTextString("承包方编码"));
					cell = row.createCell((short) 1);
					cell.setCellValue(new HSSFRichTextString("成员序号"));
					cell = row.createCell((short) 2);
					cell.setCellValue(new HSSFRichTextString("成员姓名"));
					cell = row.createCell((short) 3);
					cell.setCellValue(new HSSFRichTextString("成员性别"));
					cell = row.createCell((short) 4);
					cell.setCellValue(new HSSFRichTextString("与户主关系"));
					cell = row.createCell((short) 5);
					cell.setCellValue(new HSSFRichTextString("成员证件类型"));
					cell = row.createCell((short) 6);
					cell.setCellValue(new HSSFRichTextString("证件（公民身份证）号码"));
					cell = row.createCell((short) 7);
					cell.setCellValue(new HSSFRichTextString("是否二轮承包时的共有人"));
					cell = row.createCell((short) 8);
					cell.setCellValue(new HSSFRichTextString("成员备注代码"));
					cell = row.createCell((short) 9);
					cell.setCellValue(new HSSFRichTextString("成员备注具体说明"));

					row = jtcydcHssfSheet.createRow(1);
					cell = row.createCell((short) 0);
					cell.setCellValue(new HSSFRichTextString("CBFBM"));
					cell = row.createCell((short) 1);
					cell.setCellValue(new HSSFRichTextString("CYXH"));
					cell = row.createCell((short) 2);
					cell.setCellValue(new HSSFRichTextString("CYXM"));
					cell = row.createCell((short) 3);
					cell.setCellValue(new HSSFRichTextString("CYXB"));
					cell = row.createCell((short) 4);
					cell.setCellValue(new HSSFRichTextString("YHZGX"));
					cell = row.createCell((short) 5);
					cell.setCellValue(new HSSFRichTextString("CYZJLX"));
					cell = row.createCell((short) 6);
					cell.setCellValue(new HSSFRichTextString("CYZJHM"));
					cell = row.createCell((short) 7);
					cell.setCellValue(new HSSFRichTextString("SFGYR"));
					cell = row.createCell((short) 8);
					cell.setCellValue(new HSSFRichTextString("CYBZ"));
					cell = row.createCell((short) 9);
					cell.setCellValue(new HSSFRichTextString("CYBZSM"));

					rowNum = 2;

					row = jtcydcHssfSheet.createRow(rowNum);
					cell = row.createCell((short) 0);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("CBFBM")
							+ zm + zeroPadder(cbfbm + "", 4)));
					cell = row.createCell((short) 1);
					cell.setCellValue(new HSSFRichTextString(cyxh + ""));
					cell = row.createCell((short) 2);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("XM"))));
					// 第一行不是户主的情况
					if (!"户主".equals(StringUtil.removeNull(map.get("CYGX")))) {
						cell.setCellStyle(cellStyle);
					}
					cell = row.createCell((short) 3);
					if ("户主".equals(StringUtil.removeNull(map.get("CYGX")))) {
						cell.setCellValue(new HSSFRichTextString("男"));
						hzNum++;

						// 户主姓名不一致的情况
						if (!StringUtil.removeNull(map.get("HZ")).equals(
								StringUtil.removeNull(map.get("XM")))) {
							row.getCell(2).setCellStyle(cellStyle);
						}

						// 户主备注为死亡的情况
						if (StringUtil.removeNull(map.get("BZ")).indexOf("死") != -1
								|| StringUtil.removeNull(map.get("BZ"))
										.indexOf("亡") != -1
								|| StringUtil.removeNull(map.get("BZSM"))
										.indexOf("死") != -1
								|| StringUtil.removeNull(map.get("BZSM"))
										.indexOf("亡") != -1) {
							row.getCell(2).setCellStyle(cellStyle);
						}
					} else if (StringUtil.removeNull(map.get("CYGX")).indexOf(
							"妻") != -1
							|| StringUtil.removeNull(map.get("CYGX")).indexOf(
									"母") != -1
							|| StringUtil.removeNull(map.get("CYGX")).indexOf(
									"女") != -1
							|| StringUtil.removeNull(map.get("CYGX")).indexOf(
									"媳") != -1) {
						cell.setCellValue(new HSSFRichTextString("女"));
					} else {
						cell.setCellValue(new HSSFRichTextString("男"));
					}
					cell = row.createCell((short) 4);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("CYGX"))));
					cell = row.createCell((short) 5);
					cell.setCellValue(new HSSFRichTextString("居民身份证"));
					cell = row.createCell((short) 6);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("SFZ"))));
					cell = row.createCell((short) 7);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("GYR"))));
					cell = row.createCell((short) 8);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("BZ"))));
					cell = row.createCell((short) 9);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("BZSM"))));

					// 获取户主姓名和身份证号码
					if (!"".equals(StringUtil.removeNull(map.get("HZ")))) {
						hz = StringUtil.removeNull(map.get("HZ"));
						hzsfz = StringUtil.removeNull(map.get("SFZ"));
					}
					// 获取每户的土地使用证编号
					if (!"".equals(StringUtil.removeNull(map.get("TDSYZ")))) {
						if (StringUtil.removeNull(map.get("TDSYZ"))
								.indexOf(".") != -1) {
							tdsyzbh = StringUtil.removeNull(map.get("TDSYZ"))
									.substring(
											0,
											StringUtil.removeNull(
													map.get("TDSYZ")).indexOf(
													"."));
						} else {
							tdsyzbh = StringUtil.removeNull(map.get("TDSYZ"));
						}
					}

					// 创建承包方调查sheet
					cbfdcHssfSheet = wb.createSheet("承包方调查表");

					row = cbfdcHssfSheet.createRow(0);
					cell = row.createCell((short) 0);
					cell.setCellValue(new HSSFRichTextString("是否闭户"));
					cell = row.createCell((short) 1);
					cell.setCellValue(new HSSFRichTextString("承包方编码"));
					cell = row.createCell((short) 2);
					cell.setCellValue(new HSSFRichTextString("承包方名称"));
					cell = row.createCell((short) 3);
					cell.setCellValue(new HSSFRichTextString("承包方类型"));
					cell = row.createCell((short) 4);
					cell.setCellValue(new HSSFRichTextString("承包方证件类型"));
					cell = row.createCell((short) 5);
					cell.setCellValue(new HSSFRichTextString("承包方证件号码"));
					cell = row.createCell((short) 6);
					cell.setCellValue(new HSSFRichTextString("联系电话"));
					cell = row.createCell((short) 7);
					cell.setCellValue(new HSSFRichTextString("承包方地址"));
					cell = row.createCell((short) 8);
					cell.setCellValue(new HSSFRichTextString("邮政编码"));
					cell = row.createCell((short) 9);
					cell.setCellValue(new HSSFRichTextString("承包方成员数量"));
					cell = row.createCell((short) 10);
					cell.setCellValue(new HSSFRichTextString("原经营权证号"));
					cell = row.createCell((short) 11);
					cell.setCellValue(new HSSFRichTextString("原承包合同号"));
					cell = row.createCell((short) 12);
					cell.setCellValue(new HSSFRichTextString("承包起始日期"));
					cell = row.createCell((short) 13);
					cell.setCellValue(new HSSFRichTextString("承包终止日期"));
					cell = row.createCell((short) 14);
					cell.setCellValue(new HSSFRichTextString("调查员"));
					cell = row.createCell((short) 15);
					cell.setCellValue(new HSSFRichTextString("调查日期"));
					cell = row.createCell((short) 16);
					cell.setCellValue(new HSSFRichTextString("调查记事"));
					cell = row.createCell((short) 17);
					cell.setCellValue(new HSSFRichTextString("公示记事"));
					cell = row.createCell((short) 18);
					cell.setCellValue(new HSSFRichTextString("公示记事人"));
					cell = row.createCell((short) 19);
					cell.setCellValue(new HSSFRichTextString("公示记事日期"));

					row = cbfdcHssfSheet.createRow(1);
					cell = row.createCell((short) 0);
					cell.setCellValue(new HSSFRichTextString("SFBH"));
					cell = row.createCell((short) 1);
					cell.setCellValue(new HSSFRichTextString("CBFBM"));
					cell = row.createCell((short) 2);
					cell.setCellValue(new HSSFRichTextString("CBFMC"));
					cell = row.createCell((short) 3);
					cell.setCellValue(new HSSFRichTextString("CBFLX"));
					cell = row.createCell((short) 4);
					cell.setCellValue(new HSSFRichTextString("CBFZJLX"));
					cell = row.createCell((short) 5);
					cell.setCellValue(new HSSFRichTextString("CBFZJHM"));
					cell = row.createCell((short) 6);
					cell.setCellValue(new HSSFRichTextString("LXDH"));
					cell = row.createCell((short) 7);
					cell.setCellValue(new HSSFRichTextString("CBFDZ"));
					cell = row.createCell((short) 8);
					cell.setCellValue(new HSSFRichTextString("YZBM"));
					cell = row.createCell((short) 9);
					cell.setCellValue(new HSSFRichTextString("CBFCYSL"));
					cell = row.createCell((short) 10);
					cell.setCellValue(new HSSFRichTextString("YCBQZBH"));
					cell = row.createCell((short) 11);
					cell.setCellValue(new HSSFRichTextString("YCBHTBH"));
					cell = row.createCell((short) 12);
					cell.setCellValue(new HSSFRichTextString("YCBQSRQ"));
					cell = row.createCell((short) 13);
					cell.setCellValue(new HSSFRichTextString("YCBZZRQ"));
					cell = row.createCell((short) 14);
					cell.setCellValue(new HSSFRichTextString("CBFDCY"));
					cell = row.createCell((short) 15);
					cell.setCellValue(new HSSFRichTextString("CBFDCRQ"));
					cell = row.createCell((short) 16);
					cell.setCellValue(new HSSFRichTextString("CBFDCJS"));
					cell = row.createCell((short) 17);
					cell.setCellValue(new HSSFRichTextString("GSJS"));
					cell = row.createCell((short) 18);
					cell.setCellValue(new HSSFRichTextString("GSJSR"));
					cell = row.createCell((short) 19);
					cell.setCellValue(new HSSFRichTextString("GSJSRQ"));

					cbfRowNum = 1;

					// 创建地块调查sheet
					dkdcHssfSheet = wb.createSheet("地块调查表");

					row = dkdcHssfSheet.createRow(0);
					cell = row.createCell((short) 0);
					cell.setCellValue(new HSSFRichTextString("承包方标准编码（18位）"));
					cell = row.createCell((short) 1);
					cell.setCellValue(new HSSFRichTextString("承包方代表"));
					cell = row.createCell((short) 2);
					cell.setCellValue(new HSSFRichTextString("承包权取得方式"));
					cell = row.createCell((short) 3);
					cell.setCellValue(new HSSFRichTextString("所在大田标识"));
					cell = row.createCell((short) 4);
					cell.setCellValue(new HSSFRichTextString("地块连接号"));
					cell = row.createCell((short) 5);
					cell.setCellValue(new HSSFRichTextString("地块名称"));
					cell = row.createCell((short) 6);
					cell.setCellValue(new HSSFRichTextString("地块类别"));
					cell = row.createCell((short) 7);
					cell.setCellValue(new HSSFRichTextString("应享园地面积"));
					cell = row.createCell((short) 8);
					cell.setCellValue(new HSSFRichTextString("地块长度"));
					cell = row.createCell((short) 9);
					cell.setCellValue(new HSSFRichTextString("地块宽度"));
					cell = row.createCell((short) 10);
					cell.setCellValue(new HSSFRichTextString("地块丈量面积"));
					cell = row.createCell((short) 11);
					cell.setCellValue(new HSSFRichTextString("是否宅基地包园丈地块"));
					cell = row.createCell((short) 12);
					cell.setCellValue(new HSSFRichTextString(
							"包园丈合同面积（亩），即地块的承包田面积部分"));
					cell = row.createCell((short) 13);
					cell.setCellValue(new HSSFRichTextString(
							"已减扣的应享三地（园地、自留地、饲料地）面积（亩），即地块的园地面积部分"));
					cell = row.createCell((short) 14);
					cell.setCellValue(new HSSFRichTextString("原始二轮承包合同面积（亩）"));
					cell = row.createCell((short) 15);
					cell.setCellValue(new HSSFRichTextString("实地已征占的合同面积（亩）"));
					cell = row.createCell((short) 16);
					cell.setCellValue(new HSSFRichTextString("归算后的合同面积（亩）"));
					cell = row.createCell((short) 17);
					cell.setCellValue(new HSSFRichTextString("地块东至"));
					cell = row.createCell((short) 18);
					cell.setCellValue(new HSSFRichTextString("地块南至"));
					cell = row.createCell((short) 19);
					cell.setCellValue(new HSSFRichTextString("地块西至"));
					cell = row.createCell((short) 20);
					cell.setCellValue(new HSSFRichTextString("地块北至"));
					cell = row.createCell((short) 21);
					cell.setCellValue(new HSSFRichTextString("所有权性质"));
					cell = row.createCell((short) 22);
					cell.setCellValue(new HSSFRichTextString("是否基本农田"));
					cell = row.createCell((short) 23);
					cell.setCellValue(new HSSFRichTextString("地力等级"));
					cell = row.createCell((short) 24);
					cell.setCellValue(new HSSFRichTextString("土地用途"));
					cell = row.createCell((short) 25);
					cell.setCellValue(new HSSFRichTextString("土地利用类型"));
					cell = row.createCell((short) 26);
					cell.setCellValue(new HSSFRichTextString("地块备注信息"));

					row = dkdcHssfSheet.createRow(1);
					cell = row.createCell((short) 0);
					cell.setCellValue(new HSSFRichTextString("CBFBM"));
					cell = row.createCell((short) 1);
					cell.setCellValue(new HSSFRichTextString("CBFMC"));
					cell = row.createCell((short) 2);
					cell.setCellValue(new HSSFRichTextString("CBQQDFS"));
					cell = row.createCell((short) 3);
					cell.setCellValue(new HSSFRichTextString("DTBS"));
					cell = row.createCell((short) 4);
					cell.setCellValue(new HSSFRichTextString("DKLJH"));
					cell = row.createCell((short) 5);
					cell.setCellValue(new HSSFRichTextString("DKMC"));
					cell = row.createCell((short) 6);
					cell.setCellValue(new HSSFRichTextString("DKLB"));
					cell = row.createCell((short) 7);
					cell.setCellValue(new HSSFRichTextString("R_YXYDMJ"));
					cell = row.createCell((short) 8);
					cell.setCellValue(new HSSFRichTextString("DKCD"));
					cell = row.createCell((short) 9);
					cell.setCellValue(new HSSFRichTextString("DKKD"));
					cell = row.createCell((short) 10);
					cell.setCellValue(new HSSFRichTextString("DKZLMJ"));
					cell = row.createCell((short) 11);
					cell.setCellValue(new HSSFRichTextString("R_SFBYZZJD"));
					cell = row.createCell((short) 12);
					cell.setCellValue(new HSSFRichTextString("R_BYZHTMJ"));
					cell = row.createCell((short) 13);
					cell.setCellValue(new HSSFRichTextString("R_BYZJKSDMJ"));
					cell = row.createCell((short) 14);
					cell.setCellValue(new HSSFRichTextString("OriHTMJM"));
					cell = row.createCell((short) 15);
					cell.setCellValue(new HSSFRichTextString("ZZHTMJM"));
					cell = row.createCell((short) 16);
					cell.setCellValue(new HSSFRichTextString("HTMJ"));
					cell = row.createCell((short) 17);
					cell.setCellValue(new HSSFRichTextString("DKDZ"));
					cell = row.createCell((short) 18);
					cell.setCellValue(new HSSFRichTextString("DKNZ"));
					cell = row.createCell((short) 19);
					cell.setCellValue(new HSSFRichTextString("DKXZ"));
					cell = row.createCell((short) 20);
					cell.setCellValue(new HSSFRichTextString("DKBZ"));
					cell = row.createCell((short) 21);
					cell.setCellValue(new HSSFRichTextString("SYQXZ"));
					cell = row.createCell((short) 22);
					cell.setCellValue(new HSSFRichTextString("SFJBNT"));
					cell = row.createCell((short) 23);
					cell.setCellValue(new HSSFRichTextString("DLDJ"));
					cell = row.createCell((short) 24);
					cell.setCellValue(new HSSFRichTextString("TDYT"));
					cell = row.createCell((short) 25);
					cell.setCellValue(new HSSFRichTextString("TDLYLX"));
					cell = row.createCell((short) 26);
					cell.setCellValue(new HSSFRichTextString("DKBZXX"));

					dkRowNum = 2;

					row = dkdcHssfSheet.createRow(dkRowNum);
					cell = row.createCell((short) 0);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("CBFBM")
							+ zm + zeroPadder(cbfbm + "", 4)));
					cell = row.createCell((short) 1);
					cell.setCellValue(new HSSFRichTextString(hz));
					cell = row.createCell((short) 2);
					cell.setCellValue(new HSSFRichTextString("家庭承包"));
					cell = row.createCell((short) 3);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 4);
					cell.setCellValue(new HSSFRichTextString(cbfbm + "-"
							+ dkljh));
					cell = row.createCell((short) 5);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("DM"))));
					cell = row.createCell((short) 6);
					cell.setCellValue(new HSSFRichTextString("承包地块"));
					cell = row.createCell((short) 7);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("ZJD"))));
					cell = row.createCell((short) 8);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("DC"))));
					cell = row.createCell((short) 9);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("DK"))));
					cell = row.createCell((short) 10);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("ZRMJ"))));
					cell = row.createCell((short) 11);
					if ("宅基地".equals(StringUtil.removeNull(map.get("DM")))) {
						cell.setCellValue(new HSSFRichTextString("是"));
					} else {
						cell.setCellValue(new HSSFRichTextString("否"));
					}
					cell = row.createCell((short) 12);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("ZSMJ"))));
					cell = row.createCell((short) 13);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 14);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 15);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 16);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 17);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("DD"))));
					cell = row.createCell((short) 18);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("DN"))));
					cell = row.createCell((short) 19);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("DX"))));
					cell = row.createCell((short) 20);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("DB"))));
					cell = row.createCell((short) 21);
					cell.setCellValue(new HSSFRichTextString("30 | 集体土地所有权"));
					cell = row.createCell((short) 22);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 23);
					cell.setCellValue(new HSSFRichTextString("00 | 未定等"));
					cell = row.createCell((short) 24);
					if ("宅基地".equals(StringUtil.removeNull(map.get("DM")))) {
						cell.setCellValue(new HSSFRichTextString("5 | 非农业用途"));
					} else {
						cell.setCellValue(new HSSFRichTextString("1 | 种植业"));
					}
					cell = row.createCell((short) 25);
					if ("宅基地".equals(StringUtil.removeNull(map.get("DM")))) {
						cell.setCellValue(new HSSFRichTextString("072"));
					} else {
						cell.setCellValue(new HSSFRichTextString("011"));
					}
					cell = row.createCell((short) 26);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("DKBZ"))));

					if (list.size() == 1) {
						// 创建发包方记录
						row = fbfdcHssfSheet.createRow(++fbfRowNum);
						cell = row.createCell((short) 0);
						cell.setCellValue(new HSSFRichTextString(inputMap
								.get("CBFBM")
								+ zm));
						cell = row.createCell((short) 1);
						cell.setCellValue(new HSSFRichTextString(inputMap
								.get("CBFDZ")
								+ "经济合作社（" + zm + "组）"));
						cell = row.createCell((short) 2);
						cell.setCellValue(new HSSFRichTextString(inputMap
								.get("CBFDZ")
								+ "（" + zm + "组）"));
						cell = row.createCell((short) 3);
						cell.setCellValue(new HSSFRichTextString(""));
						cell = row.createCell((short) 4);
						cell.setCellValue(new HSSFRichTextString("居民身份证"));
						cell = row.createCell((short) 5);
						cell.setCellValue(new HSSFRichTextString(""));
						cell = row.createCell((short) 6);
						cell.setCellValue(new HSSFRichTextString(""));
						cell = row.createCell((short) 7);
						cell.setCellValue(new HSSFRichTextString(inputMap
								.get("CBFDZ")));
						cell = row.createCell((short) 8);
						cell.setCellValue(new HSSFRichTextString(inputMap
								.get("YZBM")));
						cell = row.createCell((short) 9);
						cell.setCellValue(new HSSFRichTextString(inputMap
								.get("DCY")));
						cell = row.createCell((short) 10);
						cell.setCellValue(new HSSFRichTextString(inputMap
								.get("DCRQ")));
						cell = row.createCell((short) 11);
						cell.setCellValue(new HSSFRichTextString(""));
						cell = row.createCell((short) 12);
						cell.setCellValue(new HSSFRichTextString(""));

						// 创建承包方记录
						row = cbfdcHssfSheet.createRow(++cbfRowNum);
						cell = row.createCell((short) 0);
						cell.setCellValue(new HSSFRichTextString("否"));
						cell = row.createCell((short) 1);
						cell.setCellValue(new HSSFRichTextString(inputMap
								.get("CBFBM")
								+ zm + zeroPadder(cbfbm + "", 4)));
						cell = row.createCell((short) 2);
						cell.setCellValue(new HSSFRichTextString(hz));
						cell = row.createCell((short) 3);
						cell.setCellValue(new HSSFRichTextString("农户"));
						cell = row.createCell((short) 4);
						cell.setCellValue(new HSSFRichTextString("居民身份证"));
						cell = row.createCell((short) 5);
						cell.setCellValue(new HSSFRichTextString(hzsfz));
						cell = row.createCell((short) 6);
						cell.setCellValue(new HSSFRichTextString(""));
						cell = row.createCell((short) 7);
						cell.setCellValue(new HSSFRichTextString(inputMap
								.get("CBFDZ")
								+ zm + "组"));
						cell = row.createCell((short) 8);
						cell.setCellValue(new HSSFRichTextString(inputMap
								.get("YZBM")));
						cell = row.createCell((short) 9);
						cell.setCellValue(new HSSFRichTextString(cyxh + ""));
						cell = row.createCell((short) 10);
						cell.setCellValue(new HSSFRichTextString(tdsyzbh));
						cell = row.createCell((short) 11);
						cell.setCellValue(new HSSFRichTextString(tdsyzbh));
						cell = row.createCell((short) 12);
						cell.setCellValue(new HSSFRichTextString("1997年9月1日"));
						cell = row.createCell((short) 13);
						cell.setCellValue(new HSSFRichTextString("2027年8月31日"));
						cell = row.createCell((short) 14);
						cell.setCellValue(new HSSFRichTextString(inputMap
								.get("DCY")));
						cell = row.createCell((short) 15);
						cell.setCellValue(new HSSFRichTextString(inputMap
								.get("DCRQ")));
						cell = row.createCell((short) 16);
						cell.setCellValue(new HSSFRichTextString(""));
						cell = row.createCell((short) 17);
						cell.setCellValue(new HSSFRichTextString(""));
						cell = row.createCell((short) 18);
						cell.setCellValue(new HSSFRichTextString(""));
						cell = row.createCell((short) 19);
						cell.setCellValue(new HSSFRichTextString(""));

						OutputStream os = new FileOutputStream(outFile);
						wb.write(os);
						os.close();
						continue;
					}
					continue;
				}

				if (!StringUtil.removeNull(map.get("ZM")).equals(
						StringUtil.removeNull(tempMap.get("ZM")))) {
					// 创建发包方记录
					row = fbfdcHssfSheet.createRow(++fbfRowNum);
					cell = row.createCell((short) 0);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("CBFBM")
							+ zm));
					cell = row.createCell((short) 1);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("CBFDZ")
							+ "经济合作社（" + zm + "组）"));
					cell = row.createCell((short) 2);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("CBFDZ")
							+ "（" + zm + "组）"));
					cell = row.createCell((short) 3);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 4);
					cell.setCellValue(new HSSFRichTextString("居民身份证"));
					cell = row.createCell((short) 5);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 6);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 7);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("CBFDZ")));
					cell = row.createCell((short) 8);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("YZBM")));
					cell = row.createCell((short) 9);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("DCY")));
					cell = row.createCell((short) 10);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("DCRQ")));
					cell = row.createCell((short) 11);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 12);
					cell.setCellValue(new HSSFRichTextString(""));

					// 创建承包方记录
					row = cbfdcHssfSheet.createRow(++cbfRowNum);
					cell = row.createCell((short) 0);
					cell.setCellValue(new HSSFRichTextString("否"));
					cell = row.createCell((short) 1);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("CBFBM")
							+ zm + zeroPadder(cbfbm + "", 4)));
					cell = row.createCell((short) 2);
					cell.setCellValue(new HSSFRichTextString(hz));
					cell = row.createCell((short) 3);
					cell.setCellValue(new HSSFRichTextString("农户"));
					cell = row.createCell((short) 4);
					cell.setCellValue(new HSSFRichTextString("居民身份证"));
					cell = row.createCell((short) 5);
					cell.setCellValue(new HSSFRichTextString(hzsfz));
					cell = row.createCell((short) 6);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 7);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("CBFDZ")
							+ zm + "组"));
					cell = row.createCell((short) 8);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("YZBM")));
					cell = row.createCell((short) 9);
					cell.setCellValue(new HSSFRichTextString(cyxh + ""));
					cell = row.createCell((short) 10);
					cell.setCellValue(new HSSFRichTextString(tdsyzbh));
					cell = row.createCell((short) 11);
					cell.setCellValue(new HSSFRichTextString(tdsyzbh));
					cell = row.createCell((short) 12);
					cell.setCellValue(new HSSFRichTextString("1997年9月1日"));
					cell = row.createCell((short) 13);
					cell.setCellValue(new HSSFRichTextString("2027年8月31日"));
					cell = row.createCell((short) 14);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("DCY")));
					cell = row.createCell((short) 15);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("DCRQ")));
					cell = row.createCell((short) 16);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 17);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 18);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 19);
					cell.setCellValue(new HSSFRichTextString(""));

					OutputStream os = new FileOutputStream(outFile);
					wb.write(os);
					os.close();

					tempMap = map;

					zm = StringUtil.removeNull(tempMap.get("ZM"));
					if (zm.indexOf("组") != -1) {
						zm = zm.substring(0, zm.indexOf("组"));
					} else if (zm.indexOf(".") != -1) {
						zm = zm.substring(0, zm.indexOf("."));
					}
					zm = zeroPadder(zm, 2);

					cbfbm = 1;
					cyxh = 1;
					hzNum = 0;
					dkljh = 1;

					outFile = new File(f + File.separator + zm + "组.xlsx");

					wb = new HSSFWorkbook();

					// 设置表格颜色
					hssfFont = wb.createFont();
					hssfFont.setColor(HSSFFont.COLOR_RED);
					cellStyle = wb.createCellStyle();
					cellStyle.setFont(hssfFont);

					// 创建发包方调查sheet
					fbfdcHssfSheet = wb.createSheet("发包方调查表");

					row = fbfdcHssfSheet.createRow(0);
					cell = row.createCell((short) 0);
					cell.setCellValue(new HSSFRichTextString("发包方编码"));
					cell = row.createCell((short) 1);
					cell.setCellValue(new HSSFRichTextString("发包方名称"));
					cell = row.createCell((short) 2);
					cell.setCellValue(new HSSFRichTextString("发包方村组行政名称"));
					cell = row.createCell((short) 3);
					cell.setCellValue(new HSSFRichTextString("发包方负责人姓名"));
					cell = row.createCell((short) 4);
					cell.setCellValue(new HSSFRichTextString("负责人证件类型"));
					cell = row.createCell((short) 5);
					cell.setCellValue(new HSSFRichTextString("负责人证件号码"));
					cell = row.createCell((short) 6);
					cell.setCellValue(new HSSFRichTextString("联系电话"));
					cell = row.createCell((short) 7);
					cell.setCellValue(new HSSFRichTextString("发包方地址"));
					cell = row.createCell((short) 8);
					cell.setCellValue(new HSSFRichTextString("邮政编码"));
					cell = row.createCell((short) 9);
					cell.setCellValue(new HSSFRichTextString("调查员"));
					cell = row.createCell((short) 10);
					cell.setCellValue(new HSSFRichTextString("调查日期"));
					cell = row.createCell((short) 11);
					cell.setCellValue(new HSSFRichTextString("调查记事"));
					cell = row.createCell((short) 12);
					cell.setCellValue(new HSSFRichTextString("发包方原负责人姓名"));

					row = fbfdcHssfSheet.createRow(1);
					cell = row.createCell((short) 0);
					cell.setCellValue(new HSSFRichTextString("FBFBM"));
					cell = row.createCell((short) 1);
					cell.setCellValue(new HSSFRichTextString("FBFMC"));
					cell = row.createCell((short) 2);
					cell.setCellValue(new HSSFRichTextString("FBFXZMC"));
					cell = row.createCell((short) 3);
					cell.setCellValue(new HSSFRichTextString("FBFFZRXM"));
					cell = row.createCell((short) 4);
					cell.setCellValue(new HSSFRichTextString("FZRZJLX"));
					cell = row.createCell((short) 5);
					cell.setCellValue(new HSSFRichTextString("FZRZJHM"));
					cell = row.createCell((short) 6);
					cell.setCellValue(new HSSFRichTextString("LXDH"));
					cell = row.createCell((short) 7);
					cell.setCellValue(new HSSFRichTextString("FBFDZ"));
					cell = row.createCell((short) 8);
					cell.setCellValue(new HSSFRichTextString("YZBM"));
					cell = row.createCell((short) 9);
					cell.setCellValue(new HSSFRichTextString("FBFDCY"));
					cell = row.createCell((short) 10);
					cell.setCellValue(new HSSFRichTextString("FBFDCRQ"));
					cell = row.createCell((short) 11);
					cell.setCellValue(new HSSFRichTextString("FBFDCJS"));
					cell = row.createCell((short) 12);
					cell.setCellValue(new HSSFRichTextString("FBFYFZRXM"));

					fbfRowNum = 1;

					// 创建家庭成员调查sheet
					jtcydcHssfSheet = wb.createSheet("家庭成员调查表");

					row = jtcydcHssfSheet.createRow(0);
					cell = row.createCell((short) 0);
					cell.setCellValue(new HSSFRichTextString("承包方编码"));
					cell = row.createCell((short) 1);
					cell.setCellValue(new HSSFRichTextString("成员序号"));
					cell = row.createCell((short) 2);
					cell.setCellValue(new HSSFRichTextString("成员姓名"));
					cell = row.createCell((short) 3);
					cell.setCellValue(new HSSFRichTextString("成员性别"));
					cell = row.createCell((short) 4);
					cell.setCellValue(new HSSFRichTextString("与户主关系"));
					cell = row.createCell((short) 5);
					cell.setCellValue(new HSSFRichTextString("成员证件类型"));
					cell = row.createCell((short) 6);
					cell.setCellValue(new HSSFRichTextString("证件（公民身份证）号码"));
					cell = row.createCell((short) 7);
					cell.setCellValue(new HSSFRichTextString("是否二轮承包时的共有人"));
					cell = row.createCell((short) 8);
					cell.setCellValue(new HSSFRichTextString("成员备注代码"));
					cell = row.createCell((short) 9);
					cell.setCellValue(new HSSFRichTextString("成员备注具体说明"));

					row = jtcydcHssfSheet.createRow(1);
					cell = row.createCell((short) 0);
					cell.setCellValue(new HSSFRichTextString("CBFBM"));
					cell = row.createCell((short) 1);
					cell.setCellValue(new HSSFRichTextString("CYXH"));
					cell = row.createCell((short) 2);
					cell.setCellValue(new HSSFRichTextString("CYXM"));
					cell = row.createCell((short) 3);
					cell.setCellValue(new HSSFRichTextString("CYXB"));
					cell = row.createCell((short) 4);
					cell.setCellValue(new HSSFRichTextString("YHZGX"));
					cell = row.createCell((short) 5);
					cell.setCellValue(new HSSFRichTextString("CYZJLX"));
					cell = row.createCell((short) 6);
					cell.setCellValue(new HSSFRichTextString("CYZJHM"));
					cell = row.createCell((short) 7);
					cell.setCellValue(new HSSFRichTextString("SFGYR"));
					cell = row.createCell((short) 8);
					cell.setCellValue(new HSSFRichTextString("CYBZ"));
					cell = row.createCell((short) 9);
					cell.setCellValue(new HSSFRichTextString("CYBZSM"));

					rowNum = 2;
					row = jtcydcHssfSheet.createRow(rowNum);
					cell = row.createCell((short) 0);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("CBFBM")
							+ zm + zeroPadder(cbfbm + "", 4)));
					cell = row.createCell((short) 1);
					cell.setCellValue(new HSSFRichTextString(cyxh + ""));
					cell = row.createCell((short) 2);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("XM"))));
					if (!"户主".equals(StringUtil.removeNull(map.get("CYGX")))) {
						cell.setCellStyle(cellStyle);
					}
					cell = row.createCell((short) 3);
					if ("户主".equals(StringUtil.removeNull(map.get("CYGX")))) {
						cell.setCellValue(new HSSFRichTextString("男"));
						hzNum++;

						if (!StringUtil.removeNull(map.get("HZ")).equals(
								StringUtil.removeNull(map.get("XM")))) {
							row.getCell(2).setCellStyle(cellStyle);
						}

						if (StringUtil.removeNull(map.get("BZ")).indexOf("死") != -1
								|| StringUtil.removeNull(map.get("BZ"))
										.indexOf("亡") != -1
								|| StringUtil.removeNull(map.get("BZSM"))
										.indexOf("死") != -1
								|| StringUtil.removeNull(map.get("BZSM"))
										.indexOf("亡") != -1) {
							row.getCell(2).setCellStyle(cellStyle);
						}
					} else if (StringUtil.removeNull(map.get("CYGX")).indexOf(
							"妻") != -1
							|| StringUtil.removeNull(map.get("CYGX")).indexOf(
									"母") != -1
							|| StringUtil.removeNull(map.get("CYGX")).indexOf(
									"女") != -1
							|| StringUtil.removeNull(map.get("CYGX")).indexOf(
									"媳") != -1) {
						cell.setCellValue(new HSSFRichTextString("女"));
					} else {
						cell.setCellValue(new HSSFRichTextString("男"));
					}
					cell = row.createCell((short) 4);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("CYGX"))));
					cell = row.createCell((short) 5);
					cell.setCellValue(new HSSFRichTextString("居民身份证"));
					cell = row.createCell((short) 6);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("SFZ"))));
					cell = row.createCell((short) 7);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("GYR"))));
					cell = row.createCell((short) 8);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("BZ"))));
					cell = row.createCell((short) 9);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("BZSM"))));

					// 获取户主姓名和身份证号码
					if (!"".equals(StringUtil.removeNull(map.get("HZ")))) {
						hz = StringUtil.removeNull(map.get("HZ"));
						hzsfz = StringUtil.removeNull(map.get("SFZ"));
					}

					// 获取每户的土地使用证编号
					if (!"".equals(StringUtil.removeNull(map.get("TDSYZ")))) {
						if (StringUtil.removeNull(map.get("TDSYZ"))
								.indexOf(".") != -1) {
							tdsyzbh = StringUtil.removeNull(map.get("TDSYZ"))
									.substring(
											0,
											StringUtil.removeNull(
													map.get("TDSYZ")).indexOf(
													"."));
						} else {
							tdsyzbh = StringUtil.removeNull(map.get("TDSYZ"));
						}
					}

					// 创建承包方调查sheet
					cbfdcHssfSheet = wb.createSheet("承包方调查表");

					row = cbfdcHssfSheet.createRow(0);
					cell = row.createCell((short) 0);
					cell.setCellValue(new HSSFRichTextString("是否闭户"));
					cell = row.createCell((short) 1);
					cell.setCellValue(new HSSFRichTextString("承包方编码"));
					cell = row.createCell((short) 2);
					cell.setCellValue(new HSSFRichTextString("承包方名称"));
					cell = row.createCell((short) 3);
					cell.setCellValue(new HSSFRichTextString("承包方类型"));
					cell = row.createCell((short) 4);
					cell.setCellValue(new HSSFRichTextString("承包方证件类型"));
					cell = row.createCell((short) 5);
					cell.setCellValue(new HSSFRichTextString("承包方证件号码"));
					cell = row.createCell((short) 6);
					cell.setCellValue(new HSSFRichTextString("联系电话"));
					cell = row.createCell((short) 7);
					cell.setCellValue(new HSSFRichTextString("承包方地址"));
					cell = row.createCell((short) 8);
					cell.setCellValue(new HSSFRichTextString("邮政编码"));
					cell = row.createCell((short) 9);
					cell.setCellValue(new HSSFRichTextString("承包方成员数量"));
					cell = row.createCell((short) 10);
					cell.setCellValue(new HSSFRichTextString("原经营权证号"));
					cell = row.createCell((short) 11);
					cell.setCellValue(new HSSFRichTextString("原承包合同号"));
					cell = row.createCell((short) 12);
					cell.setCellValue(new HSSFRichTextString("承包起始日期"));
					cell = row.createCell((short) 13);
					cell.setCellValue(new HSSFRichTextString("承包终止日期"));
					cell = row.createCell((short) 14);
					cell.setCellValue(new HSSFRichTextString("调查员"));
					cell = row.createCell((short) 15);
					cell.setCellValue(new HSSFRichTextString("调查日期"));
					cell = row.createCell((short) 16);
					cell.setCellValue(new HSSFRichTextString("调查记事"));
					cell = row.createCell((short) 17);
					cell.setCellValue(new HSSFRichTextString("公示记事"));
					cell = row.createCell((short) 18);
					cell.setCellValue(new HSSFRichTextString("公示记事人"));
					cell = row.createCell((short) 19);
					cell.setCellValue(new HSSFRichTextString("公示记事日期"));

					row = cbfdcHssfSheet.createRow(1);
					cell = row.createCell((short) 0);
					cell.setCellValue(new HSSFRichTextString("SFBH"));
					cell = row.createCell((short) 1);
					cell.setCellValue(new HSSFRichTextString("CBFBM"));
					cell = row.createCell((short) 2);
					cell.setCellValue(new HSSFRichTextString("CBFMC"));
					cell = row.createCell((short) 3);
					cell.setCellValue(new HSSFRichTextString("CBFLX"));
					cell = row.createCell((short) 4);
					cell.setCellValue(new HSSFRichTextString("CBFZJLX"));
					cell = row.createCell((short) 5);
					cell.setCellValue(new HSSFRichTextString("CBFZJHM"));
					cell = row.createCell((short) 6);
					cell.setCellValue(new HSSFRichTextString("LXDH"));
					cell = row.createCell((short) 7);
					cell.setCellValue(new HSSFRichTextString("CBFDZ"));
					cell = row.createCell((short) 8);
					cell.setCellValue(new HSSFRichTextString("YZBM"));
					cell = row.createCell((short) 9);
					cell.setCellValue(new HSSFRichTextString("CBFCYSL"));
					cell = row.createCell((short) 10);
					cell.setCellValue(new HSSFRichTextString("YCBQZBH"));
					cell = row.createCell((short) 11);
					cell.setCellValue(new HSSFRichTextString("YCBHTBH"));
					cell = row.createCell((short) 12);
					cell.setCellValue(new HSSFRichTextString("YCBQSRQ"));
					cell = row.createCell((short) 13);
					cell.setCellValue(new HSSFRichTextString("YCBZZRQ"));
					cell = row.createCell((short) 14);
					cell.setCellValue(new HSSFRichTextString("CBFDCY"));
					cell = row.createCell((short) 15);
					cell.setCellValue(new HSSFRichTextString("CBFDCRQ"));
					cell = row.createCell((short) 16);
					cell.setCellValue(new HSSFRichTextString("CBFDCJS"));
					cell = row.createCell((short) 17);
					cell.setCellValue(new HSSFRichTextString("GSJS"));
					cell = row.createCell((short) 18);
					cell.setCellValue(new HSSFRichTextString("GSJSR"));
					cell = row.createCell((short) 19);
					cell.setCellValue(new HSSFRichTextString("GSJSRQ"));

					cbfRowNum = 1;

					// 创建地块调查sheet
					dkdcHssfSheet = wb.createSheet("地块调查表");

					row = dkdcHssfSheet.createRow(0);
					cell = row.createCell((short) 0);
					cell.setCellValue(new HSSFRichTextString("承包方标准编码（18位）"));
					cell = row.createCell((short) 1);
					cell.setCellValue(new HSSFRichTextString("承包方代表"));
					cell = row.createCell((short) 2);
					cell.setCellValue(new HSSFRichTextString("承包权取得方式"));
					cell = row.createCell((short) 3);
					cell.setCellValue(new HSSFRichTextString("所在大田标识"));
					cell = row.createCell((short) 4);
					cell.setCellValue(new HSSFRichTextString("地块连接号"));
					cell = row.createCell((short) 5);
					cell.setCellValue(new HSSFRichTextString("地块名称"));
					cell = row.createCell((short) 6);
					cell.setCellValue(new HSSFRichTextString("地块类别"));
					cell = row.createCell((short) 7);
					cell.setCellValue(new HSSFRichTextString("应享园地面积"));
					cell = row.createCell((short) 8);
					cell.setCellValue(new HSSFRichTextString("地块长度"));
					cell = row.createCell((short) 9);
					cell.setCellValue(new HSSFRichTextString("地块宽度"));
					cell = row.createCell((short) 10);
					cell.setCellValue(new HSSFRichTextString("地块丈量面积"));
					cell = row.createCell((short) 11);
					cell.setCellValue(new HSSFRichTextString("是否宅基地包园丈地块"));
					cell = row.createCell((short) 12);
					cell.setCellValue(new HSSFRichTextString(
							"包园丈合同面积（亩），即地块的承包田面积部分"));
					cell = row.createCell((short) 13);
					cell.setCellValue(new HSSFRichTextString(
							"已减扣的应享三地（园地、自留地、饲料地）面积（亩），即地块的园地面积部分"));
					cell = row.createCell((short) 14);
					cell.setCellValue(new HSSFRichTextString("原始二轮承包合同面积（亩）"));
					cell = row.createCell((short) 15);
					cell.setCellValue(new HSSFRichTextString("实地已征占的合同面积（亩）"));
					cell = row.createCell((short) 16);
					cell.setCellValue(new HSSFRichTextString("归算后的合同面积（亩）"));
					cell = row.createCell((short) 17);
					cell.setCellValue(new HSSFRichTextString("地块东至"));
					cell = row.createCell((short) 18);
					cell.setCellValue(new HSSFRichTextString("地块南至"));
					cell = row.createCell((short) 19);
					cell.setCellValue(new HSSFRichTextString("地块西至"));
					cell = row.createCell((short) 20);
					cell.setCellValue(new HSSFRichTextString("地块北至"));
					cell = row.createCell((short) 21);
					cell.setCellValue(new HSSFRichTextString("所有权性质"));
					cell = row.createCell((short) 22);
					cell.setCellValue(new HSSFRichTextString("是否基本农田"));
					cell = row.createCell((short) 23);
					cell.setCellValue(new HSSFRichTextString("地力等级"));
					cell = row.createCell((short) 24);
					cell.setCellValue(new HSSFRichTextString("土地用途"));
					cell = row.createCell((short) 25);
					cell.setCellValue(new HSSFRichTextString("土地利用类型"));
					cell = row.createCell((short) 26);
					cell.setCellValue(new HSSFRichTextString("地块备注信息"));

					row = dkdcHssfSheet.createRow(1);
					cell = row.createCell((short) 0);
					cell.setCellValue(new HSSFRichTextString("CBFBM"));
					cell = row.createCell((short) 1);
					cell.setCellValue(new HSSFRichTextString("CBFMC"));
					cell = row.createCell((short) 2);
					cell.setCellValue(new HSSFRichTextString("CBQQDFS"));
					cell = row.createCell((short) 3);
					cell.setCellValue(new HSSFRichTextString("DTBS"));
					cell = row.createCell((short) 4);
					cell.setCellValue(new HSSFRichTextString("DKLJH"));
					cell = row.createCell((short) 5);
					cell.setCellValue(new HSSFRichTextString("DKMC"));
					cell = row.createCell((short) 6);
					cell.setCellValue(new HSSFRichTextString("DKLB"));
					cell = row.createCell((short) 7);
					cell.setCellValue(new HSSFRichTextString("R_YXYDMJ"));
					cell = row.createCell((short) 8);
					cell.setCellValue(new HSSFRichTextString("DKCD"));
					cell = row.createCell((short) 9);
					cell.setCellValue(new HSSFRichTextString("DKKD"));
					cell = row.createCell((short) 10);
					cell.setCellValue(new HSSFRichTextString("DKZLMJ"));
					cell = row.createCell((short) 11);
					cell.setCellValue(new HSSFRichTextString("R_SFBYZZJD"));
					cell = row.createCell((short) 12);
					cell.setCellValue(new HSSFRichTextString("R_BYZHTMJ"));
					cell = row.createCell((short) 13);
					cell.setCellValue(new HSSFRichTextString("R_BYZJKSDMJ"));
					cell = row.createCell((short) 14);
					cell.setCellValue(new HSSFRichTextString("OriHTMJM"));
					cell = row.createCell((short) 15);
					cell.setCellValue(new HSSFRichTextString("ZZHTMJM"));
					cell = row.createCell((short) 16);
					cell.setCellValue(new HSSFRichTextString("HTMJ"));
					cell = row.createCell((short) 17);
					cell.setCellValue(new HSSFRichTextString("DKDZ"));
					cell = row.createCell((short) 18);
					cell.setCellValue(new HSSFRichTextString("DKNZ"));
					cell = row.createCell((short) 19);
					cell.setCellValue(new HSSFRichTextString("DKXZ"));
					cell = row.createCell((short) 20);
					cell.setCellValue(new HSSFRichTextString("DKBZ"));
					cell = row.createCell((short) 21);
					cell.setCellValue(new HSSFRichTextString("SYQXZ"));
					cell = row.createCell((short) 22);
					cell.setCellValue(new HSSFRichTextString("SFJBNT"));
					cell = row.createCell((short) 23);
					cell.setCellValue(new HSSFRichTextString("DLDJ"));
					cell = row.createCell((short) 24);
					cell.setCellValue(new HSSFRichTextString("TDYT"));
					cell = row.createCell((short) 25);
					cell.setCellValue(new HSSFRichTextString("TDLYLX"));
					cell = row.createCell((short) 26);
					cell.setCellValue(new HSSFRichTextString("DKBZXX"));

					dkRowNum = 2;

					row = dkdcHssfSheet.createRow(dkRowNum);
					cell = row.createCell((short) 0);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("CBFBM")
							+ zm + zeroPadder(cbfbm + "", 4)));
					cell = row.createCell((short) 1);
					cell.setCellValue(new HSSFRichTextString(hz));
					cell = row.createCell((short) 2);
					cell.setCellValue(new HSSFRichTextString("家庭承包"));
					cell = row.createCell((short) 3);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 4);
					cell.setCellValue(new HSSFRichTextString(cbfbm + "-"
							+ dkljh));
					cell = row.createCell((short) 5);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("DM"))));
					cell = row.createCell((short) 6);
					cell.setCellValue(new HSSFRichTextString("承包地块"));
					cell = row.createCell((short) 7);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("ZJD"))));
					cell = row.createCell((short) 8);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("DC"))));
					cell = row.createCell((short) 9);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("DK"))));
					cell = row.createCell((short) 10);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("ZRMJ"))));
					cell = row.createCell((short) 11);
					if ("宅基地".equals(StringUtil.removeNull(map.get("DM")))) {
						cell.setCellValue(new HSSFRichTextString("是"));
					} else {
						cell.setCellValue(new HSSFRichTextString("否"));
					}
					cell = row.createCell((short) 12);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("ZSMJ"))));
					cell = row.createCell((short) 13);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 14);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 15);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 16);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 17);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("DD"))));
					cell = row.createCell((short) 18);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("DN"))));
					cell = row.createCell((short) 19);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("DX"))));
					cell = row.createCell((short) 20);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("DB"))));
					cell = row.createCell((short) 21);
					cell.setCellValue(new HSSFRichTextString("30 | 集体土地所有权"));
					cell = row.createCell((short) 22);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 23);
					cell.setCellValue(new HSSFRichTextString("00 | 未定等"));
					cell = row.createCell((short) 24);
					if ("宅基地".equals(StringUtil.removeNull(map.get("DM")))) {
						cell.setCellValue(new HSSFRichTextString("5 | 非农业用途"));
					} else {
						cell.setCellValue(new HSSFRichTextString("1 | 种植业"));
					}
					cell = row.createCell((short) 25);
					if ("宅基地".equals(StringUtil.removeNull(map.get("DM")))) {
						cell.setCellValue(new HSSFRichTextString("072"));
					} else {
						cell.setCellValue(new HSSFRichTextString("011"));
					}
					cell = row.createCell((short) 26);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("DKBZ"))));
					continue;
				}

				if (!"".equals(StringUtil.removeNull(map.get("XM")))) {
					if (!"".equals(StringUtil.removeNull(map.get("HZ")))
							&& !map.get("HZ").equals(tempMap.get("HZ"))) {
						// 创建承包方记录
						row = cbfdcHssfSheet.createRow(++cbfRowNum);
						cell = row.createCell((short) 0);
						cell.setCellValue(new HSSFRichTextString("否"));
						cell = row.createCell((short) 1);
						cell.setCellValue(new HSSFRichTextString(inputMap
								.get("CBFBM")
								+ zm + zeroPadder(cbfbm + "", 4)));
						cell = row.createCell((short) 2);
						cell.setCellValue(new HSSFRichTextString(hz));
						cell = row.createCell((short) 3);
						cell.setCellValue(new HSSFRichTextString("农户"));
						cell = row.createCell((short) 4);
						cell.setCellValue(new HSSFRichTextString("居民身份证"));
						cell = row.createCell((short) 5);
						cell.setCellValue(new HSSFRichTextString(hzsfz));
						cell = row.createCell((short) 6);
						cell.setCellValue(new HSSFRichTextString(""));
						cell = row.createCell((short) 7);
						cell.setCellValue(new HSSFRichTextString(inputMap
								.get("CBFDZ")
								+ zm + "组"));
						cell = row.createCell((short) 8);
						cell.setCellValue(new HSSFRichTextString(inputMap
								.get("YZBM")));
						cell = row.createCell((short) 9);
						cell.setCellValue(new HSSFRichTextString(cyxh + ""));
						cell = row.createCell((short) 10);
						cell.setCellValue(new HSSFRichTextString(tdsyzbh));
						cell = row.createCell((short) 11);
						cell.setCellValue(new HSSFRichTextString(tdsyzbh));
						cell = row.createCell((short) 12);
						cell.setCellValue(new HSSFRichTextString("1997年9月1日"));
						cell = row.createCell((short) 13);
						cell.setCellValue(new HSSFRichTextString("2027年8月31日"));
						cell = row.createCell((short) 14);
						cell.setCellValue(new HSSFRichTextString(inputMap
								.get("DCY")));
						cell = row.createCell((short) 15);
						cell.setCellValue(new HSSFRichTextString(inputMap
								.get("DCRQ")));
						cell = row.createCell((short) 16);
						cell.setCellValue(new HSSFRichTextString(""));
						cell = row.createCell((short) 17);
						cell.setCellValue(new HSSFRichTextString(""));
						cell = row.createCell((short) 18);
						cell.setCellValue(new HSSFRichTextString(""));
						cell = row.createCell((short) 19);
						cell.setCellValue(new HSSFRichTextString(""));

						cbfbm++;
						cyxh = 1;

						if (hzNum != 1) {
							row = jtcydcHssfSheet.getRow(rowNum);
							row.getCell(2).setCellStyle(cellStyle);
						}
						hzNum = 0;
					} else {
						cyxh++;
					}

					row = jtcydcHssfSheet.createRow(++rowNum);
					cell = row.createCell((short) 0);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("CBFBM")
							+ zm + zeroPadder(cbfbm + "", 4)));
					cell = row.createCell((short) 1);
					cell.setCellValue(new HSSFRichTextString(cyxh + ""));
					cell = row.createCell((short) 2);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("XM"))));
					if (cyxh == 1
							&& !"户主".equals(StringUtil.removeNull(map
									.get("CYGX")))) {
						cell.setCellStyle(cellStyle);
					}
					cell = row.createCell((short) 3);
					if ("户主".equals(StringUtil.removeNull(map.get("CYGX")))) {
						cell.setCellValue(new HSSFRichTextString("男"));
						hzNum++;

						if (cyxh == 1
								&& !StringUtil.removeNull(map.get("HZ"))
										.equals(
												StringUtil.removeNull(map
														.get("XM")))) {
							row.getCell(2).setCellStyle(cellStyle);
						}

						if (cyxh == 1
								&& StringUtil.removeNull(map.get("BZ"))
										.indexOf("死") != -1
								|| StringUtil.removeNull(map.get("BZ"))
										.indexOf("亡") != -1
								|| StringUtil.removeNull(map.get("BZSM"))
										.indexOf("死") != -1
								|| StringUtil.removeNull(map.get("BZSM"))
										.indexOf("亡") != -1) {
							row.getCell(2).setCellStyle(cellStyle);
						}
					} else if (StringUtil.removeNull(map.get("CYGX")).indexOf(
							"妻") != -1
							|| StringUtil.removeNull(map.get("CYGX")).indexOf(
									"母") != -1
							|| StringUtil.removeNull(map.get("CYGX")).indexOf(
									"女") != -1
							|| StringUtil.removeNull(map.get("CYGX")).indexOf(
									"媳") != -1) {
						cell.setCellValue(new HSSFRichTextString("女"));
					} else {
						cell.setCellValue(new HSSFRichTextString("男"));
					}
					cell = row.createCell((short) 4);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("CYGX"))));
					cell = row.createCell((short) 5);
					cell.setCellValue(new HSSFRichTextString("居民身份证"));
					cell = row.createCell((short) 6);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("SFZ"))));
					cell = row.createCell((short) 7);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("GYR"))));
					cell = row.createCell((short) 8);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("BZ"))));
					cell = row.createCell((short) 9);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("BZSM"))));

					// 获取户主姓名和身份证号码
					if (!"".equals(StringUtil.removeNull(map.get("HZ")))) {
						hz = StringUtil.removeNull(map.get("HZ"));
						hzsfz = StringUtil.removeNull(map.get("SFZ"));
					}

					// 获取每户的土地使用证编号
					if (!"".equals(StringUtil.removeNull(map.get("TDSYZ")))) {
						if (StringUtil.removeNull(map.get("TDSYZ"))
								.indexOf(".") != -1) {
							tdsyzbh = StringUtil.removeNull(map.get("TDSYZ"))
									.substring(
											0,
											StringUtil.removeNull(
													map.get("TDSYZ")).indexOf(
													"."));
						} else {
							tdsyzbh = StringUtil.removeNull(map.get("TDSYZ"));
						}
					}
				} else if (!"".equals(StringUtil.removeNull(map.get("HZ")))
						&& !map.get("HZ").equals(tempMap.get("HZ"))) {
					// 户主不为空，姓名为空的处理

				}

				if (!"".equals(StringUtil.removeNull(map.get("DM")))
						&& !"合计".equals(StringUtil.removeNull(map.get("DM")))) {
					if (!"".equals(StringUtil.removeNull(map.get("HZ")))
							&& !map.get("HZ").equals(tempMap.get("HZ"))) {
						dkljh = 1;
					} else {
						dkljh++;
					}

					row = dkdcHssfSheet.createRow(++dkRowNum);
					cell = row.createCell((short) 0);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("CBFBM")
							+ zm + zeroPadder(cbfbm + "", 4)));
					cell = row.createCell((short) 1);
					cell.setCellValue(new HSSFRichTextString(hz));
					cell = row.createCell((short) 2);
					cell.setCellValue(new HSSFRichTextString("家庭承包"));
					cell = row.createCell((short) 3);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 4);
					cell.setCellValue(new HSSFRichTextString(cbfbm + "-"
							+ dkljh));
					cell = row.createCell((short) 5);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("DM"))));
					cell = row.createCell((short) 6);
					cell.setCellValue(new HSSFRichTextString("承包地块"));
					cell = row.createCell((short) 7);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("ZJD"))));
					cell = row.createCell((short) 8);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("DC"))));
					cell = row.createCell((short) 9);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("DK"))));
					cell = row.createCell((short) 10);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("ZRMJ"))));
					cell = row.createCell((short) 11);
					if ("宅基地".equals(StringUtil.removeNull(map.get("DM")))) {
						cell.setCellValue(new HSSFRichTextString("是"));
					} else {
						cell.setCellValue(new HSSFRichTextString("否"));
					}
					cell = row.createCell((short) 12);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("ZSMJ"))));
					cell = row.createCell((short) 13);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 14);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 15);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 16);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 17);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("DD"))));
					cell = row.createCell((short) 18);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("DN"))));
					cell = row.createCell((short) 19);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("DX"))));
					cell = row.createCell((short) 20);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("DB"))));
					cell = row.createCell((short) 21);
					cell.setCellValue(new HSSFRichTextString("30 | 集体土地所有权"));
					cell = row.createCell((short) 22);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 23);
					cell.setCellValue(new HSSFRichTextString("00 | 未定等"));
					cell = row.createCell((short) 24);
					if ("宅基地".equals(StringUtil.removeNull(map.get("DM")))) {
						cell.setCellValue(new HSSFRichTextString("5 | 非农业用途"));
					} else {
						cell.setCellValue(new HSSFRichTextString("1 | 种植业"));
					}
					cell = row.createCell((short) 25);
					if ("宅基地".equals(StringUtil.removeNull(map.get("DM")))) {
						cell.setCellValue(new HSSFRichTextString("072"));
					} else {
						cell.setCellValue(new HSSFRichTextString("011"));
					}
					cell = row.createCell((short) 26);
					cell.setCellValue(new HSSFRichTextString(StringUtil
							.removeNull(map.get("DKBZ"))));
				}

				if (i == list.size() - 1) {
					// 创建发包方记录
					row = fbfdcHssfSheet.createRow(++fbfRowNum);
					cell = row.createCell((short) 0);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("CBFBM")
							+ zm));
					cell = row.createCell((short) 1);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("CBFDZ")
							+ "经济合作社（" + zm + "组）"));
					cell = row.createCell((short) 2);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("CBFDZ")
							+ "（" + zm + "组）"));
					cell = row.createCell((short) 3);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 4);
					cell.setCellValue(new HSSFRichTextString("居民身份证"));
					cell = row.createCell((short) 5);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 6);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 7);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("CBFDZ")));
					cell = row.createCell((short) 8);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("YZBM")));
					cell = row.createCell((short) 9);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("DCY")));
					cell = row.createCell((short) 10);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("DCRQ")));
					cell = row.createCell((short) 11);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 12);
					cell.setCellValue(new HSSFRichTextString(""));

					// 创建承包方记录
					row = cbfdcHssfSheet.createRow(++cbfRowNum);
					cell = row.createCell((short) 0);
					cell.setCellValue(new HSSFRichTextString("否"));
					cell = row.createCell((short) 1);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("CBFBM")
							+ zm + zeroPadder(cbfbm + "", 4)));
					cell = row.createCell((short) 2);
					cell.setCellValue(new HSSFRichTextString(hz));
					cell = row.createCell((short) 3);
					cell.setCellValue(new HSSFRichTextString("农户"));
					cell = row.createCell((short) 4);
					cell.setCellValue(new HSSFRichTextString("居民身份证"));
					cell = row.createCell((short) 5);
					cell.setCellValue(new HSSFRichTextString(hzsfz));
					cell = row.createCell((short) 6);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 7);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("CBFDZ")
							+ zm + "组"));
					cell = row.createCell((short) 8);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("YZBM")));
					cell = row.createCell((short) 9);
					cell.setCellValue(new HSSFRichTextString(cyxh + ""));
					cell = row.createCell((short) 10);
					cell.setCellValue(new HSSFRichTextString(tdsyzbh));
					cell = row.createCell((short) 11);
					cell.setCellValue(new HSSFRichTextString(tdsyzbh));
					cell = row.createCell((short) 12);
					cell.setCellValue(new HSSFRichTextString("1997年9月1日"));
					cell = row.createCell((short) 13);
					cell.setCellValue(new HSSFRichTextString("2027年8月31日"));
					cell = row.createCell((short) 14);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("DCY")));
					cell = row.createCell((short) 15);
					cell.setCellValue(new HSSFRichTextString(inputMap
							.get("DCRQ")));
					cell = row.createCell((short) 16);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 17);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 18);
					cell.setCellValue(new HSSFRichTextString(""));
					cell = row.createCell((short) 19);
					cell.setCellValue(new HSSFRichTextString(""));

					OutputStream os = new FileOutputStream(outFile);
					wb.write(os);
					os.close();
				}
			}

		} catch (Exception e) {
			e.printStackTrace();
			flag = 0;
		} finally {
			return flag;
		}

	}

	/**
	 * 对外提供读取excel 的方法
	 * */
	public static List<Map<String, Object>> readExcel(File file, int sheetNum,
			int maxTitleRow) throws IOException {
		String fileName = file.getName();
		String extension = fileName.lastIndexOf(".") == -1 ? "" : fileName
				.substring(fileName.lastIndexOf(".") + 1);
		if ("xls".equals(extension)) {
			return read2003Excel(file, sheetNum, maxTitleRow);
		} else if ("xlsx".equals(extension)) {
			return read2007Excel(file, sheetNum, maxTitleRow);
		} else {
			throw new IOException("不支持的文件类型");
		}
	}

	private static Object getCellValueByCellType(HSSFCell cell) {
		Object value = null;
		switch (cell.getCellType()) {
		case XSSFCell.CELL_TYPE_STRING:
			value = cell.getStringCellValue();
			break;
		case XSSFCell.CELL_TYPE_FORMULA:
			try {
				value = nf.format(cell.getNumericCellValue());
			} catch (Exception e) {
				value = String.valueOf(cell.getRichStringCellValue());
			}
			break;
		case XSSFCell.CELL_TYPE_NUMERIC:
			/*
			 * if ("@".equals(cell.getCellStyle().getDataFormatString())) {
			 * value = cell.getNumericCellValue(); } else if
			 * ("General".equals(cell.getCellStyle() .getDataFormatString())) {
			 * value = cell.getNumericCellValue(); } else { value =
			 * HSSFDateUtil.getJavaDate(cell.getNumericCellValue()); }
			 */
			value = nf.format(cell.getNumericCellValue());
			break;
		case XSSFCell.CELL_TYPE_BOOLEAN:
			value = cell.getBooleanCellValue();
			break;
		case XSSFCell.CELL_TYPE_BLANK:
			value = "";
			break;
		default:
			value = cell.toString();
		}

		return value;
	}

	/**
	 * 读取 office 2003 excel
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	private static List<Map<String, Object>> read2003Excel(File file,
			int sheetNum, int maxTitleRow) throws IOException {
		titles = new HashMap<String, Integer>();

		List<Map<String, Object>> list = new LinkedList<Map<String, Object>>();

		HSSFWorkbook hwb = new HSSFWorkbook(new FileInputStream(file));
		HSSFSheet sheet = hwb.getSheetAt(sheetNum);

		HSSFRow row = null;
		HSSFCell cell = null;
		Object value = null;

		for (int i = sheet.getFirstRowNum(); i < maxTitleRow; i++) {
			row = sheet.getRow(i);
			if (row == null) {
				continue;
			}
			for (int j = row.getFirstCellNum(); j < row.getLastCellNum(); j++) {
				cell = row.getCell(j);
				if (cell == null) {
					continue;
				}

				value = getCellValueByCellType(cell);

				if (value == null || "".equals(value)) {
					continue;
				}
				if ("组名".equals(value.toString().trim())) {
					titles.put("ZM", j);
				} else if ("户主".equals(value.toString().trim())) {
					titles.put("HZ", j);
				} else if (value.toString().trim().indexOf("身份证") != -1) {
					titles.put("XM", j - 2);
					titles.put("CYGX", j - 1);
					titles.put("SFZ", j);
				} else if (value.toString().trim().indexOf("二轮") != -1) {
					titles.put("GYR", j);
				} else if ("地块备注".equals(value.toString().trim())) {
					titles.put("DKBZ", j);
				} else if ("备注".equals(value.toString().trim())) {
					titles.put("BZ", j);
				} else if (value.toString().trim().indexOf("备注说明") != -1) {
					titles.put("BZSM", j);
				} else if (value.toString().trim().indexOf("土地使用证") != -1) {
					titles.put("TDSYZ", j);
				} else if (value.toString().trim().indexOf("地名") != -1) {
					titles.put("DM", j);
				} else if (value.toString().trim().indexOf("宅基地") != -1) {
					titles.put("ZJD", j);
				} else if (value.toString().trim().indexOf("长") != -1) {
					titles.put("DC", j);
				} else if (value.toString().trim().indexOf("宽") != -1) {
					titles.put("DK", j);
				} else if (value.toString().trim().indexOf("自然面积") != -1) {
					titles.put("ZRMJ", j);
				} else if (value.toString().trim().indexOf("折算面积") != -1) {
					titles.put("ZSMJ", j);
				} else if (value.toString().trim().indexOf("东") != -1) {
					titles.put("DD", j);
				} else if (value.toString().trim().indexOf("西") != -1) {
					titles.put("DX", j);
				} else if (value.toString().trim().indexOf("南") != -1) {
					titles.put("DN", j);
				} else if (value.toString().trim().indexOf("北") != -1) {
					titles.put("DB", j);
				}
			}
		}

		//
		for (int i = maxTitleRow; i < sheet.getPhysicalNumberOfRows(); i++) {
			row = sheet.getRow(i);
			if (row == null) {
				continue;
			}

			Map<String, Object> map = new HashMap<String, Object>();

			for (Map.Entry<String, Integer> entry : titles.entrySet()) {
				String key = entry.getKey();
				cell = row.getCell(entry.getValue());
				if (cell == null) {
					value = "";
				} else {
					value = getCellValueByCellType(cell);
				}
				map.put(key, value);
			}
			list.add(map);
		}
		return list;
	}

	private static Object getCellValueByCellType(XSSFCell cell) {
		Object value = null;
		switch (cell.getCellType()) {
		case XSSFCell.CELL_TYPE_STRING:
			value = cell.getStringCellValue();
			break;
		case XSSFCell.CELL_TYPE_FORMULA:
			try {
				value = nf.format(cell.getNumericCellValue());
			} catch (Exception e) {
				value = String.valueOf(cell.getRichStringCellValue());
			}
			break;
		case XSSFCell.CELL_TYPE_NUMERIC:
			/*
			 * if ("@".equals(cell.getCellStyle().getDataFormatString())) {
			 * value = cell.getNumericCellValue(); } else if
			 * ("General".equals(cell.getCellStyle() .getDataFormatString())) {
			 * value = cell.getNumericCellValue(); } else { value =
			 * HSSFDateUtil.getJavaDate(cell.getNumericCellValue()); }
			 */
			value = nf.format(cell.getNumericCellValue());
			break;
		case XSSFCell.CELL_TYPE_BOOLEAN:
			value = cell.getBooleanCellValue();
			break;
		case XSSFCell.CELL_TYPE_BLANK:
			value = "";
			break;
		default:
			value = cell.toString();
		}

		return value;
	}

	/**
	 * 读取Office 2007 excel
	 * */
	private static List<Map<String, Object>> read2007Excel(File file,
			int sheetNum, int maxTitleRow) throws IOException {
		titles = new HashMap<String, Integer>();

		List<Map<String, Object>> list = new LinkedList<Map<String, Object>>();

		// 构造 XSSFWorkbook 对象，strPath 传入文件路径
		XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(file));
		// 读取第一章表格内容
		XSSFSheet sheet = xwb.getSheetAt(sheetNum);

		XSSFRow row = null;
		XSSFCell cell = null;
		Object value = null;
		for (int i = sheet.getFirstRowNum(); i < maxTitleRow; i++) {
			row = sheet.getRow(i);
			if (row == null) {
				continue;
			}
			for (int j = row.getFirstCellNum(); j < row.getLastCellNum(); j++) {
				cell = row.getCell(j);
				if (cell == null) {
					continue;
				}

				value = getCellValueByCellType(cell);

				if (value == null || "".equals(value)) {
					continue;
				}
				if ("组名".equals(value.toString().trim())) {
					titles.put("ZM", j);
				} else if ("户主".equals(value.toString().trim())) {
					titles.put("HZ", j);
				} else if (value.toString().trim().indexOf("身份证") != -1) {
					titles.put("XM", j - 2);
					titles.put("CYGX", j - 1);
					titles.put("SFZ", j);
				} else if (value.toString().trim().indexOf("二轮") != -1) {
					titles.put("GYR", j);
				} else if ("地块备注".equals(value.toString().trim())) {
					titles.put("DKBZ", j);
				} else if ("备注".equals(value.toString().trim())) {
					titles.put("BZ", j);
				} else if (value.toString().trim().indexOf("备注说明") != -1) {
					titles.put("BZSM", j);
				} else if (value.toString().trim().indexOf("土地使用证") != -1) {
					titles.put("TDSYZ", j);
				} else if (value.toString().trim().indexOf("地名") != -1) {
					titles.put("DM", j);
				} else if (value.toString().trim().indexOf("宅基地") != -1) {
					titles.put("ZJD", j);
				} else if (value.toString().trim().indexOf("长") != -1) {
					titles.put("DC", j);
				} else if (value.toString().trim().indexOf("宽") != -1) {
					titles.put("DK", j);
				} else if (value.toString().trim().indexOf("自然面积") != -1) {
					titles.put("ZRMJ", j);
				} else if (value.toString().trim().indexOf("折算面积") != -1) {
					titles.put("ZSMJ", j);
				} else if (value.toString().trim().indexOf("东") != -1) {
					titles.put("DD", j);
				} else if (value.toString().trim().indexOf("西") != -1) {
					titles.put("DX", j);
				} else if (value.toString().trim().indexOf("南") != -1) {
					titles.put("DN", j);
				} else if (value.toString().trim().indexOf("北") != -1) {
					titles.put("DB", j);
				}
			}
		}

		//
		for (int i = maxTitleRow; i < sheet.getPhysicalNumberOfRows(); i++) {
			row = sheet.getRow(i);
			if (row == null) {
				continue;
			}

			Map<String, Object> map = new HashMap<String, Object>();

			for (Map.Entry<String, Integer> entry : titles.entrySet()) {
				String key = entry.getKey();
				cell = row.getCell(entry.getValue());
				if (cell == null) {
					value = "";
				} else {
					value = getCellValueByCellType(cell);
				}
				map.put(key, value);
			}
			list.add(map);
		}
		return list;
	}

	private static String zeroPadder(String s, int order) {
		if (s.length() >= order) {
			return s;
		}
		char[] data = new char[order];
		Arrays.fill(data, '0');

		for (int i = s.length() - 1, j = order - 1; i >= 0; i--, j--) {
			data[j] = s.charAt(i);
		}
		return String.valueOf(data);
	}
}