package com.kk.yyzc;

import java.math.BigDecimal;

public class BigDecimalTest {
	public static void main(String[] args) {
		BigDecimal bd1 = new BigDecimal("0.05");
		BigDecimal bd2 = new BigDecimal("0.01");
		bd1 = bd1.add(bd2);
		System.out.println(bd1);
		
		float f1 = 0.05f;
		float f2 = 0.01f;
		System.out.println(f1 + f2);
		System.out.println(System.getProperty("file.separator"));
	}
}
