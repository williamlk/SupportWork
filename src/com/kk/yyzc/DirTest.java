package com.kk.yyzc;

import java.io.FileNotFoundException;

public class DirTest {
	public static void main(String[] args) throws FileNotFoundException {
		//需要在bin文件下添加金融服务平台（个人）投产数据统计-汇总xxxx-xxxx.et
		String realPath = DirTest.class.getResource("金融服务平台（个人）投产数据统计-汇总xxxx-xxxx.et").getPath();
		java.io.File file = new java.io.File(realPath);
        realPath = file.getAbsolutePath();
        System.out.println(realPath);
        
        System.out.println(DirTest.class.getResourceAsStream("金融服务平台（个人）投产数据统计-汇总xxxx-xxxx.et")); 
	}
}
