package com.kk.yyzc;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
/**
 * 填寫个人网银、个人掌银、注册量、软token新签约的4種筆數
 * @author kk
 *
 */
public class GetDataByTxt {
	// 目标文件所在的地址、目标文件名、目标文件所要写的sheet
	static String TARGET_EXCEL_ADDRESS = "E:\\工作相关\\小林林要提数\\20161114";
	static String TARGET_EXCEL = "金融服务平台（个人）投产数据统计-汇总1110-1111.et";
	static String TARGET_EXCEL_SHEET_NAME = "重点产品交易数据";
	
	// 读取sql脚本查询出来的数据源文件地址，及源文件名
	static String SQL_ADDRESS = "E:\\工作相关\\小林林要提数\\20161114\\kk";
	static String SQL_FORBCDE = "kk_20161114.out";
	static String[] SQL_FORP = {"kk_20161114_5.out", "kk_20161114_6.out", "kk_20161114_7.out", "kk_20161114_8.out"};
	
	static String FILE_SEPARATOR = System.getProperty("file.separator");
	
	@SuppressWarnings({ "unchecked", "rawtypes"})
	public static void main(String[] args) {
		long beginTime = System.currentTimeMillis();
		try {
			// 统计SQL_FORBCDE中的数值到listBCDE中
			BufferedReader brForBCDE = new BufferedReader(new InputStreamReader(new FileInputStream(SQL_ADDRESS + FILE_SEPARATOR + SQL_FORBCDE)));
			String temp = null;
			List listBCDE = new ArrayList<String>();
			while((temp = brForBCDE.readLine()) != null) {
				try {
					if(!"".equals(temp.trim())) {
					listBCDE.add(Integer.parseInt(temp.trim()));
					}
				} catch(Exception e) {
					continue;
				}
			}
			brForBCDE.close();
			
			// 写入目标文件中的BCDE列, 故天数可以通过  总记录数/要记录的列数   算出天数
			int size = listBCDE.size();
			int days = size / 4;
			System.out.println(SQL_FORBCDE + "中共有：" + size + "条记录, 您统计的天数为：" + days + "天, 请核对");
			Workbook target = new HSSFWorkbook(new FileInputStream(TARGET_EXCEL_ADDRESS + FILE_SEPARATOR + TARGET_EXCEL));
			Sheet targetSheet = target.getSheet(TARGET_EXCEL_SHEET_NAME);
			for(int num = 0; num < days; num++) {
				for(int i = 0; i < 4; i++) {
					// 将计算出的数据写入到目标文件中的个人网银B、个人掌银C、注册量D、软token新签约E的4个列表中
					try {
						targetSheet.getRow(2 + num).createCell(1 + i).setCellValue(listBCDE.get(4 * num + i).toString());
						target.write(new FileOutputStream(TARGET_EXCEL_ADDRESS + FILE_SEPARATOR + TARGET_EXCEL));
					} catch (FileNotFoundException e) {
						System.out.println("请关闭" + TARGET_EXCEL + "在进行操作！");
						e.printStackTrace();
						System.exit(0);
					} catch (IOException e) {
						e.printStackTrace();
					}
				}
			}
			System.out.println("小林林要读数的B、C、D、E列已写入。");
			
			//统计SQL_FORP中的需要记录到listP中
			List<BigDecimal> listP = new ArrayList<BigDecimal>();
			for(String str : SQL_FORP) {
				BufferedReader brForP = new BufferedReader(new InputStreamReader(new FileInputStream(SQL_ADDRESS + FILE_SEPARATOR + str)));
				temp = null;
				while((temp = brForP.readLine()) != null) {
					try {
						if(!"".equals(temp.trim())) {
							listP.add(BigDecimal.valueOf(
									Double.parseDouble(temp.substring(temp.length() - 20, temp.length()).trim())));
						}
					} catch(Exception e) {
						continue;
					}
				}
				brForP.close();
			}
			
			//将读取的数据写入到目标文件的网银端交易额P列中
			int sizeListP = listP.size();
			System.out.println(SQL_FORP[0] + "中共有：" + sizeListP / SQL_FORP.length + "條記錄, 您统计的总记录数为：" + sizeListP + "条");
			for(int num = 0; num < days; num++) {
				// 读取辅助中第5，6，7，8数据库，故这里为4
				BigDecimal db = new BigDecimal("0.0");
				for(int i = 0; i < 4; i++) {
					db = db.add(listP.get(i * days + num));
				}
				try {
					targetSheet.getRow(2 + num).createCell(15).setCellValue(db.toString());
					target.write(new FileOutputStream(TARGET_EXCEL_ADDRESS + FILE_SEPARATOR + TARGET_EXCEL));
				} catch (FileNotFoundException e) {
					System.out.println("请关闭" + TARGET_EXCEL + "在进行操作！");
					e.printStackTrace();
					System.exit(0);
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			System.out.println("小林林要读数的P列已写入。");
		} catch (FileNotFoundException e) {
			System.out.println("文件不存在");
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		long endTime = System.currentTimeMillis();
		System.out.println("写入成功！用时：" + (endTime - beginTime) / 1000.0 + "秒");
	}
}
