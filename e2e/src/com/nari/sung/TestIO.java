package com.nari.sung;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class TestIO {
	private static Pattern p = Pattern.compile("([A-Za-z|-]?)(\\d+|[A-Za-z|-]+).?(\\d+|[A-Za-z|-])+");

	/**
	 * 读文件的方法
	 * 
	 * @param fName文件绝对路径
	 */
	public static void readFile(String fName) {
		try {
			FileInputStream fis = new FileInputStream(fName);
			int n = fis.read();// 读取下一个字节
			// 循环读写
			while (n != -1) {
				System.out.println("读到的字节是" + n);
				n = fis.read();
			}
			fis.close();// 关闭输入流
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * 写文件的方法
	 * 
	 * @param content写入的内容
	 * @throws Exception抛出异常
	 */
	public static void writeFile(String content) throws Exception {
		try {
			FileOutputStream fos = new FileOutputStream("D:\\" + System.currentTimeMillis() + ".txt", false);
			byte[] b = content.getBytes();// 得到组成字符串的字节
			fos.write(b);
			fos.close();// 关闭输出流
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}

	@SuppressWarnings("finally")
	public static int transFile(String fName) throws Exception {
		int flag = 1;
		BufferedReader bufferedReader = null;
		BufferedWriter bufferedWriter = null;
		try {
			bufferedReader = new BufferedReader(new FileReader(fName));
			bufferedWriter = new BufferedWriter(new FileWriter("D:\\" + System.currentTimeMillis() + ".txt"));
			String str = null;
			// 循环读写
			while ((str = bufferedReader.readLine()) != null) {
				List<String> list = new ArrayList<String>();
				Matcher m = p.matcher(str);
				while (m.find()) {
					list.add(m.group());
				}
				StringBuffer sb = new StringBuffer();
				if (3 == list.size()) {
					sb.append("  ");
					sb.append(list.get(0) + "   ");
					sb.append(list.get(2) + "   ");
					double d = Double.parseDouble(list.get(1)) * -1;
					sb.append(d + "   ");

				} else {
					sb.append("  ");
					for (String s : list) {
						sb.append(s + "   ");
					}
				}

				bufferedWriter.write(sb.toString());
				bufferedWriter.newLine();
				bufferedWriter.flush();
			}
			bufferedReader.close();// 关闭输入流
			bufferedWriter.close();
		} catch (Exception e) {
			e.printStackTrace();
			flag = 0;
		} finally {
			if (bufferedReader != null) {
				try {
					bufferedReader.close();
				} catch (Exception e) {
					e.printStackTrace();
				}
			}

			if (bufferedWriter != null) {
				try {
					bufferedWriter.close();
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
			return flag;
		}
	}

	public static void main(String[] args) {
		String str = "pP70008p9p99-1-1   -22616.390    15582.17";
		Pattern p = Pattern.compile("([A-Za-z|-]?)(\\d+|[A-Za-z|-]+).?(\\d+|[A-Za-z|-])+");
		Matcher m = p.matcher(str);
		while (m.find()) {
			System.out.println(m.group());
		}
	}

}
