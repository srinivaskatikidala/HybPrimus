package Utilities;

import java.io.FileInputStream;
import java.util.Properties;

public class PropertyFileUtil {
	public static String getValueForKey(String Key) throws Exception {

		Properties configProperties = new Properties();

		FileInputStream fis = new FileInputStream(
				"F:\\Project\\Automation\\Stock_AccountingMVN\\PropertiesFile\\RepositoryFile");
		configProperties.load(fis);
		
		return (String) configProperties.get(Key);

	}
}


