package com.kk.yyzc;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.MessageFormat;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 填写网银网捷贷笔数、网银端笔数、网银端笔数三个笔数[即K、O、S列]
 * @author kk
 *
 */
public class GetDataByExcel {
	// 目标文件所在的地址、目标文件名、目标文件所要写的sheet
	private static String TARGET_EXCEL_ADDRESS = "E:\\工作相关\\小林林要提数\\20161124";
	private static String TARGET_EXCEL = "金融服务平台（个人）投产数据统计-汇总1117-1123.et";
	private static String TARGET_EXCEL_SHEET_NAME = "重点产品交易数据";

	// 源文件所在的地址、源文件所要读的Sheet
	private static String SOURCE_EXCEL_ADDRESS = "E:\\工作相关\\小林林要提数\\20161124";
	private static String SOURCE_EXCEL_SHEET_NAME = "个人网银交易量明细";

	private static String SPACE = "";
	private static String key = SPACE;
	private static String value = SPACE;

	private static String FILE_SEPARATOR = System.getProperty("file.separator");
	
	/**
	 * 这里只对文件未关闭时抛出的编译时异常进行处理，其他异常不做处理
	 * @param args 控制台出入参数
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public static void main(String[] args) throws FileNotFoundException, IOException {
		long beginTime = System.currentTimeMillis();
		args = new String[] { "电子银行应用系统交易统计 2016年11月12日.xls",
				"电子银行应用系统交易统计 2016年11月13日.xls", 
				"电子银行应用系统交易统计 2016年11月14日.xls",
				"电子银行应用系统交易统计 2016年11月15日.xls", 
				"电子银行应用系统交易统计 2016年11月16日.xls"/*,
				"电子银行应用系统交易统计 2016年11月17日.xls", 
				"电子银行应用系统交易统计 2016年11月18日.xls" */};
		Workbook target = new HSSFWorkbook(new FileInputStream(TARGET_EXCEL_ADDRESS + FILE_SEPARATOR + TARGET_EXCEL));
		Sheet targetSheet = target.getSheet(TARGET_EXCEL_SHEET_NAME);
		for (int num = 0; num < args.length; num++) {
			// Get excel and its sheet to read data
			Workbook wb = new HSSFWorkbook(new FileInputStream(SOURCE_EXCEL_ADDRESS + FILE_SEPARATOR + args[num]));
			Sheet st = wb.getSheet(SOURCE_EXCEL_SHEET_NAME);
			Map<String, String> map = new LinkedHashMap<String, String>();

			int minRowIx = st.getFirstRowNum();
			int maxRowIx = st.getLastRowNum();
			for (int i = minRowIx; i <= maxRowIx; i++) {
				Iterator<Cell> it = st.getRow(i).cellIterator();
				// Just get two cell
				int j = 0;
				while (it.hasNext()) {
					Cell cell = it.next();
					cell.setCellType(Cell.CELL_TYPE_STRING);
					switch (j) {
					case 0:
						key = cell.getStringCellValue().trim().equals(SPACE) ? (i + SPACE) : cell.getStringCellValue().trim() + i;
						j++;
						break;
					case 1:
						value = cell.getStringCellValue();
						j++;
						break;
					}
				}
				map.put(key, value);
				key = SPACE;
				value = SPACE;
			}

			Set<Entry<String, String>> set = map.entrySet();
			Iterator<Entry<String, String>> itSet = set.iterator();
			if((maxRowIx - minRowIx + 1) == map.size()) {
				System.out.print(args[num] + "读取成功");
			} else {
				System.out.println("读取存在问题请检查" + args[num]);
			}

			int wangJieDai = 0;
			int kaiYiBao = 0;
			int daErCunDan = 0;
			while (itSet.hasNext()) {
				Entry<String, String> entry = itSet.next();
				String key = entry.getKey();
				String value = entry.getValue();
				if (key.contains("网捷贷")) {
					wangJieDai += Integer.parseInt(value);
				} else if (key.contains("快溢宝")) {
					kaiYiBao += Integer.parseInt(value);
				} else if (key.contains("大额存单") && !key.contains("大额存单-开户")) {
					daErCunDan += Integer.parseInt(value);
				}
			}
			System.out.println(MessageFormat.format("，其中网捷贷有{0}笔，快溢宝有{1}笔，大额存单有{2}笔", wangJieDai, kaiYiBao, daErCunDan));

			// 将计算出的数据写入到目标文件中
			try {
				targetSheet.getRow(2 + num).createCell(10).setCellValue(wangJieDai);
				targetSheet.getRow(2 + num).createCell(14).setCellValue(kaiYiBao);
				targetSheet.getRow(2 + num).createCell(18).setCellValue(daErCunDan);
				target.write(new FileOutputStream(TARGET_EXCEL_ADDRESS + FILE_SEPARATOR + TARGET_EXCEL));
			} catch (FileNotFoundException e) {
				System.out.println("请关闭" + TARGET_EXCEL + "在进行操作！");
				e.printStackTrace();
				System.exit(0);
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		long endTime = System.currentTimeMillis();
		System.out.println("写入成功！用时：" + (endTime - beginTime) / 1000.0 + "秒");
	}
}
