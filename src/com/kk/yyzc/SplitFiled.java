package com.kk.yyzc;

/**
 * 输入按照一定标志隔开的表字段头部以及一条记录，输出字段对应的记录
 * 例如：
 * 输入：test|test1|test2     1|2|3
 * 输出：
 * test--->1
 * test1--->2
 * test2--->3
 * @author kk
 *
 */
public class SplitFiled {
	public static void main(String[] args) {
		String files = "PROCOD|ACTNO|TRDAT|JRNNO|SEQNO|TRDAT|OWNBK|PRDNO|FRNJRN|TRTM|VCHNO|TRPROCCOD|TRBK|TRCOD|TRAMT|CSHTFR|RBIND|TRAFTBAL|ERRDAT|DBKTYP|DBKPRO|BATNO|DBKNO|APROCOD|AACNO|ABS|REM|TRCHL|TRFRM|TRPLA|ECIND|PRTIND|SUP1|SUP2|CPTPRDNO|CPTCUSNAM|TRABS|TRADDRLONG|ITMNO|CUSNO|TRINTS|TRTAX|CUSLVL|RATTYP|EXECRAT|BASERAT|RATSPRD|RATFLT|BKRATTYP|FLTAMT|REM1|TRPSG|FEEFLAG|VATTXT|TMSTP";
		
		String records = "23|407001100330262|20141205|331501293|01|20141205||6228461190002015214||092501|TS080027|23|9999|0|-0000000000500000.00|1|0|0000000000006972.11||0||00|0|23|999901910000027|杞敮||EBNK||9999|0|1||| | | | | | | | | | | | | | | | | | | | | ";
 	
		String[] filesArray = files.split("\\|");
		String[] recordsArray = records.split("\\|");
		
		if(filesArray.length != recordsArray.length) {
			System.out.println("记录数与表字段数不符！");
			System.out.println("表字段数：" + filesArray.length + "  记录个数：" + recordsArray.length);
		} else {
			for(int i = 0; i < filesArray.length; i++) {
				System.out.println(filesArray[i].toLowerCase() + "--->" + recordsArray[i]);
			}
		} 
		
	}
}
 