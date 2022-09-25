# Quick info about this app

Hi! 
The above application is just POC how to use [MS Graph to send email](https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0&tabs=http) in C# application base on .NET Framwork 4.8.

# About repository

Clone this repository, setup conifig.app and fill be free to change code. I tried to describe all classes and methods in source for easier understand and use app.
More about code is on my [blog](https://citdev.pl/blog/) (in polish).

## #1 Create Azure App and Mail-enabled security group

On first place create new Azure App on [Register an application - Microsoft Azure](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/CreateApplicationBlade/quickStartType~/null/isMSAApp~/false), save application (client) ID.

Then in Manage -> API permissions add new permissions:  `Mail.Send` as `Application` and grand admin consent.

Next [create](https://admin.microsoft.com/?auth_upn=Karol%40karolkozlowski.onmicrosoft.com&source=applauncher#/addgroupwizard) or use the Mail-enabled security group, and add members.

> **!**  Only these members will be able to send email from the app.

## #2 Create new Application Access Policy

By using Terminal connect to [Exchange Online](https://learn.microsoft.com/en-us/graph/auth-limit-mailbox-access#configure-applicationaccesspolicy). Run the following command, replacing the arguments for **AppId**, **PolicyScopeGroupId**, and **Description**:

    New-ApplicationAccessPolicy -AppId e7e4dbfc-046f-4074-9b3b-2ae8f144f59b -PolicyScopeGroupId EvenUsers@contoso.com -AccessRight RestrictAccess -Description "Restrict this app to members of distribution group EvenUsers."
- **AppId**: your Azure App
- **PolicyScopeGroupId**: your Mail-enabled security group address

> **!** Changes to application access policies can take longer than 1 hour to take effect in Microsoft Graph REST API calls

## #3 Create new Credential for Azure App

Back to your Azure App and in Manage -> Certificates & secrets, upload certificate or add a client secret (copy value of client secret).
You can create new certificate files by using this PowerShell script: [Create-SelfSignedCertificate.ps1](https://github.com/KarolFilipKozlowski/CSharp---Send-email-via-MS-Graph/blob/main/Create-SelfSignedCertificate.ps1)

## #4 Set config in CSharp - Send email via MSGraph

The last step is setup AppSecret.config files:

    <appSettings>
    	<!-- Azure App MS Graph config: -->
    	<add key = "tenantId" value = "00000000-0000-0000-0000-000000000000" />
    	<add key = "clientId" value = "00000000-0000-0000-0000-000000000000" />
    	<!-- byClientSecret: -->
    	<add key = "clientSecret" value = "--qwertyuiopasdfghjklzxcvbnm0123456789--" />
    	<!-- byCertificatePath: -->
    	<add key = "ertificatePath" value = "*.pfx" />
    	<add key = "certificatePass" value = "password" />
    	<!-- byCertificateThumbprint: -->
    	<add key = "certificateThumbp1rint" value = "QWERTYUIOPASDFGHJKLZXCVBNM0123456789QWER" />
    </appSettings>

Setup this config depending on what type of authorization you have chosen for the application.

### About me

If you like my work, consider visiting mine blog: [Karol Kozłowski | CitDev](https://citdev.pl/blog/)
