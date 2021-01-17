# Introduction 
This is a dotnet core project to connect to Microsoft Exchange Online (Microsoft 365) using EWS (Exchange Web Services) with OAuth2 and read emails.
Exchange Web Services (EWS) is a standard to enable client applications to communicate with the Exchange server. EWS provides access to much of the same data that is made available through Microsoft Office Outlook. SOAP provides the messaging framework for messages sent between the client application and the Exchange server. The SOAP messages are sent over HTTP.

# Background
Microsoft has retired support for basic authentication in EWS to connect to exchange online servers. Basic authentication here means client application passing the username and password with every request. The reason behind ending support for basic authentication is to make the communication more secure by adopting modern authentication ways (OAuth 2.0).
This has forced EWS clients using basic authentication to shift to oauth2 in order to connect to Exchange Online.

Note: Use OAuth authentication in all your new or existing EWS applications to connect to Exchange Online. OAuth authentication for EWS is only available in Exchange Online as part of Microsoft 365. EWS applications that use OAuth must be registered with Azure Active Directory first.

This project demonstrate how a background service connects to exchange online using EWS with oauth2.

# Getting Started
We are going to use OAuth authentication service provided by Azure Active Directory to enable EWS Managed applications to access Exchange Online in Microsoft 365. To use OAuth with our application we will need to:
1.	Register application with Azure Active Directory.
2.	Configure application for app-only authentication.
3.	Add code to get an authentication token.
4.	Add an authentication token to EWS requests.

# Register application with Azure Active Directory
1. To use OAuth, an application must be registeded in azure active directory first. Since ours is a console application, so we need to register it as a public client with Azure Active Directory.
2. Login to azure portal and select Azure Active Direcotry from left panel.
   ![Screenshot1](images/Screenshot1.PNG)