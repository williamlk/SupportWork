package com.kk.yyzc;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStreamReader;

import org.junit.Test;

public class ExcelMoveTest {
	@Test
	public void testMoveExcelToDestop() throws IOException {
		BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream("金融服务平台（个人）投产数据统计-汇总xxxx-xxxx.et")));
		String str = "";
		while((str = br.readLine()) != null) {
			System.out.println(str);
		}
	}
	
}
