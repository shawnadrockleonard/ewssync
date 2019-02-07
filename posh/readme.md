# Azure AD registration and Service Bus configuration instructions

This sample provides information and constructive details for synchronizing mailboxes with an On-Premises SQL Server instance.
While this is a Proof of Concept (POC), it can be used to idealize your scenario.

## Azure AD app registration steps

The below screenshots will assist in creating an App registration that can be used for secure credentials and avoid impersonation.

### Create a new application

![alt text](https://raw.githubusercontent.com/shawnadrockleonard/ewssync/master/posh/imgs/app-reg-01-list.png "Enterprise Application - listing")\
The highlighted application was created as a 'web application' type app.
You do not need to specify a public facing URL, however, it does require a reply URL to ensure you can acquire a token.
\
\
\
![alt text](https://raw.githubusercontent.com/shawnadrockleonard/ewssync/master/posh/imgs/app-reg-02-properties.png "Enterprise Application - metadata and property review")\
The application does not require additional configuration.
The screenshot of properties includes the defaults.
If you choose to enable additional capabilities, please review the information/highights.
\
\
\
![alt text](https://raw.githubusercontent.com/shawnadrockleonard/ewssync/master/posh/imgs/app-reg-03-conditional-access.png "Enterprise Application - Conditional access")\
This step is (optional).  This is great for reporting and auditing.  You can assign any number of owners who can have conditional access to the application to review, authorize, and audit the registered application.
\
\
\
![alt text](https://raw.githubusercontent.com/shawnadrockleonard/ewssync/master/posh/imgs/app-reg-04-grant-consent.png "Enterprise Application - Granting consent")\
This step depends on the purpose of the application.  If your application is like a DAEMON, in that it must run behind the scenes to generate reporting, then you will not want to require each user to consent to the application.  In this particular example, the application must read calendars.  Notice in this example that I have not enabled writing or authoring messages.  This application has the minimal permissions required to sync calendar appointments.  Once the application has the appropriate authorizations, use the 'Grant admin consent for Microsoft' to enable the appropriate authorizations.
\
\
\
![alt text](https://raw.githubusercontent.com/shawnadrockleonard/ewssync/master/posh/imgs/app-reg-05-app-registrations.png "Enterprise Application - navigate to app registrations")\
\
\
\
![alt text](https://raw.githubusercontent.com/shawnadrockleonard/ewssync/master/posh/imgs/app-reg-06-review-permissions.png "Enterprise Application - assign and review permissions")\
On this screen you can add, modify, or remove permissions.  Note: you should ensure you have minimal permissions to acheive your purpose.  If you decide to change authorizations/permissions, then you will need to navigate back to the 'Grant Consent' application screen to ensure the application has the 'new' permissions.
\
\
\
![alt text](https://raw.githubusercontent.com/shawnadrockleonard/ewssync/master/posh/imgs/app-reg-07-review-certs.png "Enterprise Application - Certificate authentication")\
We will use a self-signed certificate to enable authentication via certificates.  We need to generate a .CER file or a base64 information file that enables a 'trust' between the DAEMON and Azure AD; thus authorizing the application to execute Exchange Web Service queries.

``` posh

$cert=New-SelfSignedCertificate -Subject "CN=EWSResourceSync" -CertStoreLocation "Cert:\CurrentUser\My"  -KeyExportPolicy Exportable -KeySpec Signature
Export-PfxCertificate -Cert $cert -Password (Get-Credential).Password -FilePath .\temp.pfx -Verbose
Export-Certificate -Cert $cert -FilePath .\temp.cer -Type CERT

```

Once you have the temp.cer file on disk, you can use the 'Upload Public Key' button to add the self-signed certificate.
\
\
\
![alt text](https://raw.githubusercontent.com/shawnadrockleonard/ewssync/master/posh/imgs/app-reg-08-review-manifest.png "Enterprise Application - review manifest")\
Note: in the previous powershell statement the generated certificate is installed in the local computer store.  This is important to note as wherever you run your code/script it must be installed.  This will enable your client application to communicate with Azure to establish a trust.  Now that you've uploaded a certificate with a public thumbprint, you can review the Azure AD Application manifest.  The highlighted section indicates the unique ID of the certificate in Azure and the start/end date of the authentication.  You must take this into consideration as every certificate or password has an expiration.

## Documentation and Caveats

Exchange Web Services are amazing.  Exchange is a huge application and typically an application that incurs the highest visibility.  It scales out at a huge footprint.  The below links will help in establishing proper connectivity to avoid throttling, avoid high latency, and to ensure you gain the best response.

- How to create appointments and meetings by using EWS in Exchange 2013 <https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/>

- Notification subscriptions mailbox events and EWS in Exchange <https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services>

- EWS Streaming sample code <https://archive.codeplex.com/?p=ewsstreaming>

- When streaming fails - handle rehydration <https://docs.microsoft.com/en-us/previous-versions/office/developer/exchange-server-2010/hh312849(v%3Dexchg.140)>

- Sychronizing Events <https://docs.microsoft.com/en-us/previous-versions/office/developer/exchange-server-2010/ee693003(v%3Dexchg.80)>
