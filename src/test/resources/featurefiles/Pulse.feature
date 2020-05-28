Feature: Pulse feature test

  @pulse
  Scenario: verify for the message posted in pulse functionality
    Given Steps for pulse in case of failure
    Given Open the browser and navigate to the url
    When I enter username as "csinboundagent"
    And I enter password as "rules"
    And I click on Login button
    Then Login should be successful
    When I click on Pulse Link
    And I enter "test message" in pulse textbox
    And I click on post button
    Then "test message" should be posted successfully
    When I click on down arrow
    And I click on Logout
    Then Logout should be successful
