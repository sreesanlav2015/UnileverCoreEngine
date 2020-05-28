Feature: Login feature test

  #Author U.Ramakrishna
  @RegressionKato @SmokeKato @TestKato1
  Scenario: Login with valid credentials
    Given Steps for login in case of failure
    Given Open the browser and navigate to the url
    When I enter username as "abhinash.kotikalapudi@Areteanstech.com"
    And I enter password as "Pega1234$"
    And I click on Login button
    Then Login should be successful
