package practice_package;

import org.testng.annotations.Test;

public class Sample2Test {

	
	@Test (groups="regression")
	public void est3() {
		System.out.println("from test-3");
	}
	
	@Test(groups ="smoke")
	public void test4() {
		System.out.println("from test-4");
	}
}
