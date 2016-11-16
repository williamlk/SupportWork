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

public class GetDataByExcel {
	// Ŀ���ļ����ڵĵ�ַ��Ŀ���ļ�����Ŀ���ļ���Ҫд��sheet
	static String TARGET_EXCEL_ADDRESS = "E:\\�������\\С����Ҫ����\\20161114";
	static String TARGET_EXCEL = "���ڷ���ƽ̨�����ˣ�Ͷ������ͳ��-����1110-1111.et";
	static String TARGET_EXCEL_SHEET_NAME = "�ص��Ʒ��������";

	// Դ�ļ����ڵĵ�ַ��Դ�ļ���Ҫ����Sheet
	static String SOURCE_EXCEL_ADDRESS = "E:\\�������\\С����Ҫ����\\20161114";
	static String SOURCE_EXCEL_SHEET_NAME = "����������������ϸ";

	static String SPACE = "";
	static String key = SPACE;
	static String value = SPACE;

	static String FILE_SEPARATOR = System.getProperty("file.separator");
	
	/**
	 * @param args
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public static void main(String[] args) throws FileNotFoundException, IOException {
		long beginTime = System.currentTimeMillis();
		args = new String[] { "��������Ӧ��ϵͳ����ͳ�� 2016��11��09��.xls",
				"��������Ӧ��ϵͳ����ͳ�� 2016��11��10��.xls", 
				"��������Ӧ��ϵͳ����ͳ�� 2016��11��11��.xls"
				/*,"��������Ӧ��ϵͳ����ͳ�� 2016��11��01��.xls", 
				"��������Ӧ��ϵͳ����ͳ�� 2016��11��02��.xls",
				"��������Ӧ��ϵͳ����ͳ�� 2016��11��03��.xls", 
				"��������Ӧ��ϵͳ����ͳ�� 2016��11��04��.xls" */};
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
				System.out.print(args[num] + "��ȡ�ɹ�");
			} else {
				System.out.println("��ȡ������������" + args[num]);
			}

			int wangJieDai = 0;
			int kaiYiBao = 0;
			int daErCunDan = 0;
			while (itSet.hasNext()) {
				Entry<String, String> entry = itSet.next();
				String key = entry.getKey();
				String value = entry.getValue();
				if (key.contains("���ݴ�")) {
					wangJieDai += Integer.parseInt(value);
				} else if (key.contains("���籦")) {
					kaiYiBao += Integer.parseInt(value);
				} else if (key.contains("���浥") && !key.contains("���浥-����")) {
					daErCunDan += Integer.parseInt(value);
				}
			}
			System.out.println(MessageFormat.format("���������ݴ���{0}�ʣ����籦��{1}�ʣ����浥��{2}��", wangJieDai, kaiYiBao, daErCunDan));

			// �������������д�뵽Ŀ���ļ���
			try {
				targetSheet.getRow(2 + num).createCell(10).setCellValue(wangJieDai);
				targetSheet.getRow(2 + num).createCell(14).setCellValue(kaiYiBao);
				targetSheet.getRow(2 + num).createCell(18).setCellValue(daErCunDan);
				target.write(new FileOutputStream(TARGET_EXCEL_ADDRESS + FILE_SEPARATOR + TARGET_EXCEL));
			} catch (FileNotFoundException e) {
				System.out.println("��ر�" + TARGET_EXCEL + "�ڽ��в�����");
				e.printStackTrace();
				System.exit(0);
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		long endTime = System.currentTimeMillis();
		System.out.println("д��ɹ�����ʱ��" + (endTime - beginTime) / 1000.0 + "��");
	}
}
