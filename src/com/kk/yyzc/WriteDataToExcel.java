package com.kk.yyzc;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.math.BigDecimal;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 填写网银网捷贷笔数、网银端笔数、网银端笔数三个笔数[即K、O、S列]
 * 个人网银、个人掌银、注册量、软token新签约的4种笔数[即BCDE列]
 * 以及快e宝中网银端交易额[即P列]
 * @author kk
 *
 */
public class WriteDataToExcel {
	// 目标文件的地址（通过createExcelToProjectDir方法获取）、目标文件所要写的sheet
	private static String TARGET_EXCEL_ADDRESS = "";
	private static String TARGET_EXCEL_SHEET_NAME = "重点产品交易数据";

	// 源文件所要读的Sheet
	private static String SOURCE_EXCEL_SHEET_NAME = "个人网银交易量明细";

	private static String SPACE = "";
	private static String key = SPACE;
	private static String value = SPACE;

	private static String FILE_SEPARATOR = System.getProperty("file.separator");
	private static String USER_DIR = System.getProperty("user.dir");
	private static String project_dir = USER_DIR.split("SupportWork")[0];
	
	/**
	 * 验证并获取需要统计的电子银行应用系统交易统计 201X年XX月XX日.xls 的名称
	 * @return 返回需要统计的《电子银行应用系统交易统计 201X年XX月XX日.xls》的名称
	 * @throws IOException
	 */
	private static String[] getSourceExcelNamesForExcelKOS() {
		System.out.println("请确认电子银行应用系统交易统计 201X年XX月XX日.xls已放入" + project_dir + "中：(Y/N)");
		BufferedReader br = new BufferedReader(new InputStreamReader(System.in));
		try {
			if(!"Y".equals(br.readLine().trim().toUpperCase())) {
				System.out.println("请按照约定后再次执行本程序");
				System.exit(0);
			}
		} catch (IOException e) {
			e.printStackTrace();
			System.out.println("有问题找堃堃");
		}
		
		File files = new File(project_dir);
		String[] fileStrings = getFilterFileNames(files, "电子银行应用系统交易统计");
		if(fileStrings == null || fileStrings.length == 0) {
			System.out.println(project_dir + "目录下没有要统计的<电子银行应用系统交易统计 201X年XX月XX日.xls>");
			System.exit(0);
		}
		return fileStrings;
	}
	
	/**
	 * 获取指定文件下符合fitlerName命名的所有文件
	 * @param file 指定文件夹
	 * @param filterName 需要的文件名
	 * @return 升序排列的需要文件名所有名称的数组
	 */
	private static String[] getFilterFileNames(File file, final String filterName) {
		String[] fileStrings = file.list(new FilenameFilter() {
			
			@Override
			public boolean accept(File dir, String name) {
				return name.startsWith(filterName) || name.endsWith(filterName);
			}
		});
		Arrays.sort(fileStrings);
		return fileStrings;
	}
	
	/**
	 * 返回给定字符串的月日信息
	 * @param fileString 文件名这里文件名称必需为:电子银行应用系统交易统计 201X年XX月XX日.xls
	 * @return 返回字符串中的月日信息，例如XX月xx日
	 */
	private static String getMonthAndDay(String fileString) {
		String monthAndDay = SPACE;
		Pattern p = Pattern.compile("[0-9]+");
		Matcher m = p.matcher(fileString.split("年")[1]);
		
		while(m.find()){
		   monthAndDay += m.group() + ",";
		}
		String[] monthDay = monthAndDay.split(",");
		return monthDay[0] + "月" + monthDay[1] + "日";
	}
	
	/**
	 * 创建金融服务平台（个人）投产数据统计-汇总xxxx-xxxx.et到项目jar包所在的地址，
	 * 此方法需要在执行填写此excel信息之前执行
	 * @param startDate 开始时间
	 * @param endDate 结束时间
	 */
	private static void createExcelToProjectDir(String startDate, String endDate) {
		InputStream fis = null;
		FileOutputStream fos = null;
		try {
			TARGET_EXCEL_ADDRESS = project_dir + "\\金融服务平台（个人）投产数据统计-汇总" 
					+ startDate + "-" + endDate + ".et";
			fis = WriteDataToExcel.class.getResourceAsStream("金融服务平台（个人）投产数据统计-汇总xxxx-xxxx.et");
			//System.out.println(DirTest.class.getResource("金融服务平台（个人）投产数据统计-汇总xxxx-xxxx.et").getPath());
			
			fos = new FileOutputStream(new File(TARGET_EXCEL_ADDRESS));
			byte[] b = new byte[1024];
			while(fis.read(b) != -1) {
				fos.write(b);
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
			System.out.println("有问题找堃堃");
		} finally {
			try {
				if (fis != null) { fis.close(); }
				if (fos != null) { fos.close(); }
			} catch (IOException e) {
				e.printStackTrace();
				System.out.println("有问题找堃堃");
			}
		}
	}
	
	/**
	 * 这里只对文件未关闭时抛出的编译时异常进行处理，其他异常不做处理
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public static void main(String[] arg) throws FileNotFoundException, IOException {	
		String[] fileStrings = getSourceExcelNamesForExcelKOS();
		createExcelToProjectDir(getMonthAndDay(fileStrings[0]), getMonthAndDay(fileStrings[fileStrings.length - 1]));
		writeDataToExcelKOS(fileStrings);
		System.out.println("\n********************华丽的分割线********************\n");
		verifyAndWriteDataToExcelBCDEP();
		System.out.println("写入成功excel保存在" + project_dir);
	}

	/**
	 *  填写电子银行应用系统交易统计 201X年XX月XX日.xls中
	 *  网银网捷贷笔数、网银端笔数、网银端笔数三个笔数[即K、O、S列]
	 * @param fileStrings
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	private static void writeDataToExcelKOS(String[] fileStrings)
			throws IOException, FileNotFoundException {
		long beginTime = System.currentTimeMillis();
		Workbook target = new HSSFWorkbook(new FileInputStream(TARGET_EXCEL_ADDRESS));
		Sheet targetSheet = target.getSheet(TARGET_EXCEL_SHEET_NAME);
		for (int num = 0; num < fileStrings.length; num++) {
			// Get excel and its sheet to read data
			Workbook wb = new HSSFWorkbook(new FileInputStream(project_dir + FILE_SEPARATOR + fileStrings[num]));
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
				System.out.print(fileStrings[num] + "读取成功");
			} else {
				System.out.println("读取存在问题请检查" + fileStrings[num]);
				System.out.println("有问题找堃堃");
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
				targetSheet.getRow(2 + num).createCell(0).setCellValue(getMonthAndDay(fileStrings[num])); 
				targetSheet.getRow(2 + num).createCell(10).setCellValue(wangJieDai);
				targetSheet.getRow(2 + num).createCell(14).setCellValue(kaiYiBao);
				targetSheet.getRow(2 + num).createCell(18).setCellValue(daErCunDan);
				target.write(new FileOutputStream(TARGET_EXCEL_ADDRESS));
				
			} catch (FileNotFoundException e) {
				System.out.println("请关闭" + TARGET_EXCEL_ADDRESS + "在进行操作！");
				e.printStackTrace();
				System.exit(0);
			} catch (IOException e) {
				e.printStackTrace();
				System.out.println("有问题找堃堃");
				System.exit(0);
			}
		}
		long endTime = System.currentTimeMillis();
		System.out.println("写入K、O、S列成功！用时：" + (endTime - beginTime) / 1000.0 + "秒");
	}
	
	/**
	 * 验证123.113及102.23上读取的文件是否符合规则，
	 * 并将其中数据写入到Excel中的BCDEP列上
	 * @throws IOException
	 */
	private static void verifyAndWriteDataToExcelBCDEP() throws IOException {
		System.out.println("请确认目录" + project_dir + "下\n" +
				"1、 123.113生产机上辅助库读数的文件以_113.out结尾\n" +
				"2、 102.23上的5|6|7|8数据库对应的结尾为_5.out、_6.out、_7.out、_8.out：(Y/N)");
		BufferedReader br = new BufferedReader(new InputStreamReader(System.in));
		File files = new File(project_dir);
		String[] sql_forBCDE_array = getFilterFileNames(files, "_113.out");
		String[] sql_forP_5 = getFilterFileNames(files, "_5.out");
		String[] sql_forP_6 = getFilterFileNames(files, "_6.out");
		String[] sql_forP_7 = getFilterFileNames(files, "_7.out");
		String[] sql_forP_8 = getFilterFileNames(files, "_8.out");
		while(!"Y".equals(br.readLine().trim().toUpperCase()) || sql_forBCDE_array.length != 1 || sql_forP_5.length != 1 ||
			     sql_forP_6.length != 1 || sql_forP_7.length != 1 || sql_forP_8.length != 1) {
			System.out.println("请按照约定后输入Y");
			sql_forBCDE_array = getFilterFileNames(files, "_113.out");
			sql_forP_5 = getFilterFileNames(files, "_5.out");
			sql_forP_6 = getFilterFileNames(files, "_6.out");
			sql_forP_7 = getFilterFileNames(files, "_7.out");
			sql_forP_8 = getFilterFileNames(files, "_8.out");
		}
		
		writeDataToExcelBCDEP(sql_forBCDE_array[0], 
				new String[]{sql_forP_5[0], sql_forP_6[0], sql_forP_7[0], sql_forP_8[0]});
	}
	
	/**
	 * 填写电子银行应用系统交易统计 201X年XX月XX日.xls中
	 * 个人网银、个人掌银、注册量、软token新签约的4种笔数[即BCDE列]
	 * 以及快e宝中网银端交易额[即P列]
	 */
	private static void writeDataToExcelBCDEP(String sql_forBCDE, String[] sql_forP) {
		try {
			long beginTime = System.currentTimeMillis();
			// 统计SQL_FORBCDE中的数值到listBCDE中
			BufferedReader brForBCDE = new BufferedReader(new InputStreamReader(new FileInputStream(project_dir + FILE_SEPARATOR + sql_forBCDE)));
			String temp = null;
			List<Integer> listBCDE = new ArrayList<Integer>();
			while((temp = brForBCDE.readLine()) != null) {
				try {
					if(!SPACE.equals(temp.trim())) {
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
			System.out.println(sql_forBCDE + "中共有：" + size + "条记录, 您统计的天数为：" + days + "天, 请核对");
			Workbook target = new HSSFWorkbook(new FileInputStream(TARGET_EXCEL_ADDRESS));
			Sheet targetSheet = target.getSheet(TARGET_EXCEL_SHEET_NAME);
			for(int num = 0; num < days; num++) {
				for(int i = 0; i < 4; i++) {
					// 将计算出的数据写入到目标文件中的个人网银B、个人掌银C、注册量D、软token新签约E的4个列表中
					try {
						targetSheet.getRow(2 + num).createCell(1 + i).setCellValue(listBCDE.get(4 * num + i).toString());
						target.write(new FileOutputStream(TARGET_EXCEL_ADDRESS));
					} catch (FileNotFoundException e) {
						System.out.println("请关闭" + TARGET_EXCEL_ADDRESS + "在进行操作！");
						e.printStackTrace();
						System.exit(0);
					} catch (IOException e) {
						e.printStackTrace();
						System.out.println("有问题找堃堃");
					}
				}
			}
			System.out.println("小林林要读数的B、C、D、E列已写入。");
			
			//统计SQL_FORP中的需要记录到listP中
			List<BigDecimal> listP = new ArrayList<BigDecimal>();
			for(String str : sql_forP) {
				BufferedReader brForP = new BufferedReader(new InputStreamReader(new FileInputStream(project_dir + FILE_SEPARATOR + str)));
				temp = null;
				while((temp = brForP.readLine()) != null) {
					try {
						if(!SPACE.equals(temp.trim())) {
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
			System.out.println(sql_forP[0] + "中共有：" + sizeListP / sql_forP.length + "條記錄, 您统计的总记录数为：" + sizeListP + "条");
			for(int num = 0; num < days; num++) {
				// 读取辅助中第5，6，7，8数据库，故这里为4
				BigDecimal db = new BigDecimal("0.0");
				for(int i = 0; i < 4; i++) {
					db = db.add(listP.get(i * days + num));
				}
				try {
					targetSheet.getRow(2 + num).createCell(15).setCellValue(db.toString());
					target.write(new FileOutputStream(TARGET_EXCEL_ADDRESS));
				} catch (FileNotFoundException e) {
					System.out.println("请关闭" + TARGET_EXCEL_ADDRESS + "在进行操作！");
					e.printStackTrace();
					System.exit(0);
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			System.out.println("小林林要读数的P列已写入。");
			long endTime = System.currentTimeMillis();
			System.out.println("写入成功！用时：" + (endTime - beginTime) / 1000.0 + "秒");
		} catch (FileNotFoundException e) {
			System.out.println("文件不存在");
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
			System.out.println("有问题找堃堃");
		}
	}
}
