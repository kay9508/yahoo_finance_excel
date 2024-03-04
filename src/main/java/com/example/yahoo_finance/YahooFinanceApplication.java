package com.example.yahoo_finance;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.autoconfigure.jdbc.DataSourceAutoConfiguration;

@SpringBootApplication(
		scanBasePackages = "com.example.yahoo_finance",
				exclude={DataSourceAutoConfiguration.class} //DB를 사용하지 않고 SpringBootApplication을 실행하고 싶을 때 사용하는 옵션
		)
public class YahooFinanceApplication {

	public static void main(String[] args) {
		SpringApplication.run(YahooFinanceApplication.class, args);
	}

}
