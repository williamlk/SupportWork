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
 * ���������������������ע��������token��ǩԼ��4�N�P��
 * @author kk
 *
 */
public class GetDataByTxt {
	// Ŀ���ļ����ڵĵ�ַ��Ŀ���ļ�����Ŀ���ļ���Ҫд��sheet
	static String TARGET_EXCEL_ADDRESS = "E:\\�������\\С����Ҫ����\\20161114";
	static String TARGET_EXCEL = "���ڷ���ƽ̨�����ˣ�Ͷ������ͳ��-����1110-1111.et";
	static String TARGET_EXCEL_SHEET_NAME = "�ص��Ʒ��������";
	
	// ��ȡsql�ű���ѯ����������Դ�ļ���ַ����Դ�ļ���
	static String SQL_ADDRESS = "E:\\�������\\С����Ҫ����\\20161114\\kk";
	static String SQL_FORBCDE = "kk_20161114.out";
	static String[] SQL_FORP = {"kk_20161114_5.out", "kk_20161114_6.out", "kk_20161114_7.out", "kk_20161114_8.out"};
	
	static String FILE_SEPARATOR = System.getProperty("file.separator");
	
	@SuppressWarnings({ "unchecked", "rawtypes"})
	public static void main(String[] args) {
		long beginTime = System.currentTimeMillis();
		try {
			// ͳ��SQL_FORBCDE�е���ֵ��listBCDE��
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
			
			// д��Ŀ���ļ��е�BCDE��, ����������ͨ��  �ܼ�¼��/Ҫ��¼������   �������
			int size = listBCDE.size();
			int days = size / 4;
			System.out.println(SQL_FORBCDE + "�й��У�" + size + "����¼, ��ͳ�Ƶ�����Ϊ��" + days + "��, ��˶�");
			Workbook target = new HSSFWorkbook(new FileInputStream(TARGET_EXCEL_ADDRESS + FILE_SEPARATOR + TARGET_EXCEL));
			Sheet targetSheet = target.getSheet(TARGET_EXCEL_SHEET_NAME);
			for(int num = 0; num < days; num++) {
				for(int i = 0; i < 4; i++) {
					// �������������д�뵽Ŀ���ļ��еĸ�������B����������C��ע����D����token��ǩԼE��4���б���
					try {
						targetSheet.getRow(2 + num).createCell(1 + i).setCellValue(listBCDE.get(4 * num + i).toString());
						target.write(new FileOutputStream(TARGET_EXCEL_ADDRESS + FILE_SEPARATOR + TARGET_EXCEL));
					} catch (FileNotFoundException e) {
						System.out.println("��ر�" + TARGET_EXCEL + "�ڽ��в�����");
						e.printStackTrace();
						System.exit(0);
					} catch (IOException e) {
						e.printStackTrace();
					}
				}
			}
			System.out.println("С����Ҫ������B��C��D��E����д�롣");
			
			//ͳ��SQL_FORP�е���Ҫ��¼��listP��
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
			
			//����ȡ������д�뵽Ŀ���ļ��������˽��׶�P����
			int sizeListP = listP.size();
			System.out.println(SQL_FORP[0] + "�й��У�" + sizeListP / SQL_FORP.length + "�lӛ�, ��ͳ�Ƶ��ܼ�¼��Ϊ��" + sizeListP + "��");
			for(int num = 0; num < days; num++) {
				// ��ȡ�����е�5��6��7��8���ݿ⣬������Ϊ4
				BigDecimal db = new BigDecimal("0.0");
				for(int i = 0; i < 4; i++) {
					db = db.add(listP.get(i * days + num));
				}
				try {
					targetSheet.getRow(2 + num).createCell(15).setCellValue(db.toString());
					target.write(new FileOutputStream(TARGET_EXCEL_ADDRESS + FILE_SEPARATOR + TARGET_EXCEL));
				} catch (FileNotFoundException e) {
					System.out.println("��ر�" + TARGET_EXCEL + "�ڽ��в�����");
					e.printStackTrace();
					System.exit(0);
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			System.out.println("С����Ҫ������P����д�롣");
		} catch (FileNotFoundException e) {
			System.out.println("�ļ�������");
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		long endTime = System.currentTimeMillis();
		System.out.println("д��ɹ�����ʱ��" + (endTime - beginTime) / 1000.0 + "��");
	}
}
