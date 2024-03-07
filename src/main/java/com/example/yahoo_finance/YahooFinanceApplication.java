package com.example.yahoo_finance;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.autoconfigure.jdbc.DataSourceAutoConfiguration;

import javax.swing.*;
import java.lang.reflect.Method;

@SpringBootApplication(
		scanBasePackages = "com.example.yahoo_finance",
				exclude={DataSourceAutoConfiguration.class} //DB를 사용하지 않고 SpringBootApplication을 실행하고 싶을 때 사용하는 옵션
		)
public class YahooFinanceApplication {
	public static void main(String[] args) {
		SpringApplication.run(YahooFinanceApplication.class, args);

		String url = "http://localhost:4031";

        openURL(url);
	}

    public static void openURL(String url) {
        String osName = System.getProperty("os.name");
        try {
            if (osName.startsWith("Mac OS")) {
                Class fileMgr = Class.forName("com.apple.eio.FileManager");
                Method openURL = fileMgr.getDeclaredMethod("openURL",
                        new Class[] {String.class});
                openURL.invoke(null, new Object[] {url});
            }
            else if (osName.startsWith("Windows"))
                Runtime.getRuntime().exec("rundll32 url.dll,FileProtocolHandler " + url);
            else { //assume Unix or Linux
                String[] browsers = {
                        "firefox", "opera", "konqueror", "epiphany", "mozilla", "netscape" };
                String browser = null;
                for (int count = 0; count < browsers.length && browser == null; count++)
                    if (Runtime.getRuntime().exec(
                            new String[] {"which", browsers[count]}).waitFor() == 0)
                        browser = browsers[count];
                if (browser == null)
                    throw new Exception("Could not find web browser");
                else
                    Runtime.getRuntime().exec(new String[] {browser, url});
            }
        }
        catch (Exception e) {
            JOptionPane.showMessageDialog(null, "error" + ":\n" + e.getLocalizedMessage());
        }
    }

}
