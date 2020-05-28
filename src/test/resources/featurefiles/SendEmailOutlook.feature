Feature: Sendemail feature test

  #Author U.Ramakrishna
  @Regression @Smoke @Test3
  Scenario Outline: Send Email from outlook
    Given Enter "<path>" and "<excelname>" and "<sheet>" to send an email

    Examples: 
      | path                  | excelname     | sheet  |
      | D:\\UnileverO2CLatest | TestData.xlsx | Sheet1 |
