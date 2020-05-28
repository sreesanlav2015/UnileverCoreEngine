Feature: Search feature test

  @Reg
  Scenario: Search for the case
	Given Steps for search in case of failure 
	Given Open the browser and navigate to the url 
	When I enter username as "csinboundagent" 
	And I enter password as "rules" 
	And I click on Login button 
	Then Login should be successful 
	When I search for the valid caseid 
		| caseid  | Row |
		| I-18535 |   0 |
	Then case should be displayed 
	When I click on down arrow 
	And I click on Logout 
	Then Logout should be successful 
