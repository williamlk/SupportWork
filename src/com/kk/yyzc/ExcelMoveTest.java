package com.kk.yyzc;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;

import org.junit.Test;

public class ExcelMoveTest {
	@Test
	public void testMoveExcelToDestop() throws IOException {
		FileInputStream fis = new FileInputStream(new File("金融服务平台（个人）投产数据统计-汇总xxxx-xxxx.et"));
		FileOutputStream fos = new FileOutputStream(new File(System.getProperty("user.home") + "\\Desktop\\金融服务平台（个人）投产数据统计-汇总1111-1112.et"));
		byte[] b = new byte[1024];
		while(fis.read(b) != -1) {
			fos.write(b);
		}
	}
}
