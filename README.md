# MspPasswordResetNotification
Script to run on a Domain Controller that runs Active Directory in order to notify users of approaching password expiry

NOTE : This script is untested as a whole, but portions of it have been tested in different environments.

The general flow of this script is :
  Import the modules necessary to run the script
  Get the date and set how many days out to warn the user
  Get a list of users whose passwords can expire or are already expired
  For each user in the environment whose password can expire, it will get the expiry date and convert it to date time
    It will then calculate the days remaining and store that number for later use
    Generates a subject
    Grabs the user's email and generates the appropriate message
    Connects to Graph with ID's
    Specifies sender
    Specifies Body
    Puts the message in drafts
    Sends message
    
    That is for each user

In order to use this script, you will need to be able to use an email address that can be authenticated using Graph
  This address will need to be put in on line 130
  
There is information to fill out 
  Line 81 : email address that users can email with issues
  Line 81 : Phone number they can call
  Line 106 : Phone number they can call
  Line 119 : AppId to connect to Graph - the idea was to make this a variable for each customer
  Line 120 : TenantID of managed company - the idea was to make this a variable for each customer
  Line 121 : ClientSecret ID for Graph - the idea was to make this a variable for each customer
  
This code is maintained by Steve Denison 7/7/2022
