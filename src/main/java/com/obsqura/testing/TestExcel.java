package com.obsqura.testing;

import java.io.IOException;

public class TestExcel {

	
	public static void main(String[] args) throws IOException{
		
		String s=ExcelProgram.readStringData(1, 0);
		System.out.println(s);
		double d=ExcelProgram.readNumericData(1, 1);
		System.out.println(d);

	}

}
