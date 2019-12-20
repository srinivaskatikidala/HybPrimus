package DriveFactory;

import org.testng.annotations.Test;

public class AppTest{
	
	@Test
	public void kickstart() throws Throwable{
		
		try {
			DriverScript ds = new DriverScript();
			ds.startTest();
		} catch (Exception e) {
			
			e.printStackTrace();
		}
		
	}
	
}