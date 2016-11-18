package com.kk.yyzc;

import java.math.BigDecimal;

public class BigDecimalTest {
	public static void main(String[] args) {
		BigDecimal bd1 = new BigDecimal("0.05");
		BigDecimal bd2 = new BigDecimal("0.011");
		bd1 = bd1.add(bd2);
		System.out.println(bd1);
		
		float f1 = 0.05f;
		float f2 = 0.011f;
		System.out.println(f1 + f2);
		
		// 使用BigDecimal注意点不能用原生的数据，只能使用String类型来作为行参
		BigDecimal bd3 = new BigDecimal(0.05f);
		BigDecimal bd4 = new BigDecimal(0.011f);
		bd3 = bd3.add(bd4);
		System.out.println(bd3);
	}
}
