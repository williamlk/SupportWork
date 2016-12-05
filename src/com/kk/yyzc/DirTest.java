package com.kk.yyzc;

import java.io.File;
import java.io.FileNotFoundException;

public class DirTest {
	public static void main(String[] args) throws FileNotFoundException {
		//1、需要在bin文件下添加金融服务平台（个人）投产数据统计-汇总xxxx-xxxx.et，此类加载器的路径在此class文件处于的bin路径下
		String realPath = DirTest.class.getResource("金融服务平台（个人）投产数据统计-汇总xxxx-xxxx.et").getPath();
		File file = new File(realPath);
        realPath = file.getAbsolutePath();
        System.out.println(realPath);
        
        System.out.println(DirTest.class.getResourceAsStream("金融服务平台（个人）投产数据统计-汇总xxxx-xxxx.et")); 
	}
}
