package com.cucumber.framework.runner;


import cucumber.api.CucumberOptions;
import cucumber.api.testng.AbstractTestNGCucumberTests;

@CucumberOptions(
		
features = {"classpath:featurefiles",
		
		
           }, 
glue = {
		"classpath:com.cucumber.framework.stepDef",
		"classpath:com.cucumber.framework.PageObjects",
		"classpath:com.cucumber.framework.TestBase",
		"classpath:com.cucumber.framework.helper",
		}, 

plugin = {"html:target/cucumber-html-report",
		"json:target/cucumber-report/cucumber.json",
		"pretty:target/cucumber-pretty.txt",
		"junit:target/cucumber-results.xml"
		},

monochrome = true,

dryRun = false



)
public class CommonRunner extends AbstractTestNGCucumberTests {

}



