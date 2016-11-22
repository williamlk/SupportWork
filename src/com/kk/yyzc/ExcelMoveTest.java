package com.kk.yyzc;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.util.Arrays;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.junit.Test;

public class ExcelMoveTest {
	@Test
	public void testMoveExcelToDestop() {
		FileInputStream fis = null;
		FileOutputStream fos = null;
		try {
			fis = new FileInputStream(new File("金融服务平台（个人）投产数据统计-汇总xxxx-xxxx.et"));
			fos = new FileOutputStream(new File(System.getProperty("user.home") + "\\Desktop\\金融服务平台（个人）投产数据统计-汇总1111-1112.et"));
			byte[] b = new byte[1024];
			while(fis.read(b) != -1) {
				fos.write(b);
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				if (fis != null) { fis.close(); }
				if (fos != null) { fos.close(); }
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
	
	@Test
	public void testSystemProperty() {
		String dir = System.getProperty("user.dir");
		System.out.println(dir.substring(0, dir.length() - "SupportWork".length()));
		File files = new File(dir.substring(0, dir.length() - "SupportWork".length()));
		String[] fielStrings = files.list();
		for(String file : fielStrings) {
			System.out.println(file);
		}
	}
	
	@Test
	public void testFilterName() {
		File file = new File("D:/java_home");
		String[] fileStrings = file.list(new FilenameFilter() {
			
			@Override
			public boolean accept(File dir, String name) {
				return name.startsWith("电子银行应用系统交易统计");
			}
		});
		
		Arrays.sort(fileStrings);
		for(String fileString : fileStrings) {
			System.out.println(fileString);
		}
		
		System.out.println(getMonthAndDay(fileStrings[0]));
	}

	public String getMonthAndDay(String fileString) {
		String monthAndDay = "";
		Pattern p = Pattern.compile("[0-9]+");
		Matcher m = p.matcher(fileString.split("年")[1]);
		
		while(m.find()){
		   monthAndDay += m.group() + ",";
		}
		String[] monthDay = monthAndDay.split(",");
		return monthDay[0] + "月" + monthDay[1] + "日";
	}
}
