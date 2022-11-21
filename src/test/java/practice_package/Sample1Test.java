package practice_package;

import org.testng.annotations.Test;

public class Sample1Test {

	
	@Test (groups ="smoke")
	public void test1() {
		System.out.println("from test-1");
	}
	
	@Test(groups="regression")
	public void test2() {
		System.out.println("from test-2");
	}
}
