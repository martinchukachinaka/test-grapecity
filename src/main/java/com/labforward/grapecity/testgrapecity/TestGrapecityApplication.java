package com.labforward.grapecity.testgrapecity;

import com.grapecity.documents.excel.Workbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class TestGrapecityApplication implements CommandLineRunner {

	@Value("${grape-city.license}")
	private String grapeCityLicense;

	public static void main(String[] args) {
		SpringApplication.run(TestGrapecityApplication.class, args);
	}

	@Override
	public void run(String... args) {
		System.out.println("grapeCityLicense: " + grapeCityLicense);
		Workbook.SetLicenseKey(grapeCityLicense);
	}
}
