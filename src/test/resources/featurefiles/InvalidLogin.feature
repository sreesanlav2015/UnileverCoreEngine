Feature: Invalid Login feature test

  #Author U.Ramakrishna
  @Smoke
Scenario Outline: Login with Invalid credentials 
	Given Steps for login with invalid credentials "<username>" and "<password>" in case of failure 
	Given Open the browser and navigate to the url 
	When I enter username as "<username>" 
	And I enter password as "<password>" 
	And I click on Login button 
	Then User should not be logged in 
	
	Examples: 
		| username       | password |
		| Rama           | rules    |
		#| csinboundagent | rules    |
