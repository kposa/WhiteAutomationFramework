Feature: WordRelatedTests
	
Scenario: Word_01_Verify Word menu is displayed in MS Word
	Given I have installed Word
	When I launched new MS Word document
	Then I should see Word menu

Scenario: Word_02_Verify Word menu ribbon is displayed in MS Word
	Given I launched new MS Word document
	When I Click Word menu in MS Word Document
	Then I should see Word ribbon with below options
	| Menu Options         |
	| Search               |

Scenario: Word_03_Verify login screen in Word AddIn
	Given I launched new MS Word document
	When I click Login button
	Then I should be see login screen

Scenario: Word_04_Verify Cancel button in login screen in Word
	Given I launched new MS Word document
	When I see login window
	And I click Cancel button in login screen in Word
	Then Login Dialog should be disappeared

Scenario: Word_05_Verify Login in Word
	Given I launched new MS Word document
	And I Click Word menu in MS Word Document
	When I enter "stangudu" and "User@123" and click Login button
		And I selected customer "Testing"
	Then I should be logged in successfully
