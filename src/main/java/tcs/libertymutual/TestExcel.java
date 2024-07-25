package tcs.libertymutual;

import java.io.IOException;

public class TestExcel {

	public static void main(String[] args) throws IOException{
		// TODO Auto-generated method stub
		String s = ExcelCode.readStringData(1,0);
		System.out.println(s);
		 double d = ExcelCode.readNumericData(1,1);
		 System.out.println(d);

	}

}
