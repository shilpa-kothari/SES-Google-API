// Copyright 2018 Google LLC
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

// [START drive_quickstart]
package SESGoogleSheet;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;



import java.io.IOException;
import java.security.GeneralSecurityException;
import java.time.Duration;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import java.io.File;

import services.GoogleDriveService;

/* class to demonstrate use of Drive files list API */
public class GoogleDrive {

	public static void main(String... args) throws IOException, GeneralSecurityException, InterruptedException {
		String baseDir = System.getProperty("user.dir");
		String downloadDir = Paths.get(baseDir, "download").toString();
		System.out.println("Base directory: " + baseDir);
		deleteFiles(downloadDir);
		

		  System.setProperty("webdriver.chrome.driver", "C:/Users/shilpa.gadiya/Downloads/chromedriver-win64/chromedriver.exe");

		  	
	      ChromeOptions options = new ChromeOptions();
	      options.addArguments("--no-sandbox");
	      options.addArguments("--disable-dev-shm-usage");
	      options.addArguments("--disable-popup-blocking");
	      options.addArguments("--disable-extensions");
	      options.addArguments("--disable-notifications");
	      options.addArguments("--disable-infobars");
	      options.addArguments("--disable-gpu");
	      options.addArguments("--disable-translate");
	     options.addArguments("--headless");
	      options.addArguments("--window-size=1920,1080");
	      options.setExperimentalOption("prefs", getChromePreferences());
	      

//	      // Create a WebDriver instance for Chrome
	      WebDriver driver = new ChromeDriver(options);
	      WebDriverWait wait = new WebDriverWait(driver,Duration.ofSeconds(40));

	      // Navigate to the page where the file download link/button is located
	      driver.get("http://ses2.e-connectsolutions.com/");
	      
	driver.manage().window().maximize();
	      
	      driver.findElement(By.id("login_loginId")).sendKeys("102963");
	      driver.findElement(By.id("login_password")).sendKeys("default");
	      Thread.sleep(3000); 
	      driver.findElement(By.xpath("//button[@type='submit']")).click();


	      
	      Thread.sleep(3000); 
	      
	      
	      driver.findElement(By.xpath("//span[contains(@aria-label,'funnel-plot')]")).click();
	      
	      
	      wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("projectName"))).sendKeys("Raj ERP 3.0");


	      wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@class='ant-select-item-option-content']"))).click();
	      wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//span[@aria-label='search']"))).click();
	      
	      

	      Thread.sleep(5000);
	      wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("body > div:nth-child(1) > section:nth-child(1) > div:nth-child(1) > div:nth-child(1) > section:nth-child(2) > div:nth-child(1) > main:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(1) > div:nth-child(1) > div:nth-child(4) > div:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > div:nth-child(1) > div:nth-child(1) > div:nth-child(1) > div:nth-child(1) > div:nth-child(1) > table:nth-child(1) > tbody:nth-child(3) > tr:nth-child(1) > td:nth-child(5)"))).click();

	      Thread.sleep(5000);
	      // Find the download link/button
	      
	      WebElement downloadButton =  wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"rc-tabs-0-panel-ticketDetailsPending-1undefinedundefineda\"]/div/div[2]/div/div/div/div/div/div[2]/button[1]")));

	      // Click on the download link/button
	      downloadButton.click();
	      
	      Thread.sleep(8000);

		// Copy the file : temp copy file from other folder, need to comment after adding download file code
		Path sourcePath = Paths.get(baseDir, "download", "RERP3_Ticket_Details.xls");
		Path destinationPath = Paths.get(downloadDir, sourcePath.getFileName().toString());
		Files.copy(sourcePath, destinationPath);

		String filePath = getDownloadFile(downloadDir);
		GoogleDriveServicenew googleDriveService = new GoogleDriveServicenew();
		String fileId = googleDriveService.uploadSheetFileToDrive(filePath);
		GoogleSheetService.replaceSheetData(fileId);
		GoogleDrive.deleteFiles(fileId);
		

	}

	private static String getDownloadFile(String folderPath) {
		// Create a File object representing the folder
		File folder = new File(folderPath);
		String filePath = "";
		// Check if the provided path exists and is a directory
		if (folder.exists() && folder.isDirectory()) {
			// List files in the folder
			File[] files = folder.listFiles();

			// Iterate over the files and print their names
			if (files != null) {
				for (File file : files) {
					// Print file name
					System.out.println("File Name: " + file.getName());
					filePath = file.getPath();
					break;
				}
			}
		}
		return filePath;
	}

	private static void deleteFiles(String folderPath) {
		File folder = new File(folderPath);
		// List files in the folder
		File[] files = folder.listFiles();

		// Iterate over the files and delete them
		if (files != null) {
			for (File file : files) {
				if (file.isDirectory()) {
					// Recursively delete subdirectories
					deleteFiles(file.getPath());
				} else {
					// Delete the file
					if (file.delete()) {
						System.out.println("File deleted: " + file.getName());
					} else {
						System.out.println("Failed to delete file: " + file.getName());
					}
				}
			}
		}
	}
	private static java.util.HashMap<String, Object> getChromePreferences() {
	    java.util.HashMap<String, Object> chromePrefs = new java.util.HashMap<>();
	    chromePrefs.put("download.default_directory", "C:\\Users\\shilpa.gadiya\\eclipse-workspace\\SES\\download");
	    chromePrefs.put("download.prompt_for_download", false);
	    return chromePrefs;
	}


}
// [END drive_quickstart]