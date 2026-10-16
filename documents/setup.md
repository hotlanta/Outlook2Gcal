Google Calendar API Setup (One-Time Configuration)

This only needs to be completed once per user.

Ignore any scary messages about having to pay.

1. Create a Google Cloud Project
* Sign in to Google Cloud Console -> https://console.cloud.google.com/
* Click Create Project -> under IAM&Admin in the navigation panel on the left.
* Enter a name such as:	Outlook2Gcal
* Click Create.
2. Enable the Google Calendar API
* Open the newly created project.
* Navigate to: APIs & Services → Library -> https://console.cloud.google.com/apis/dashboard?project=calendar-api-510113
* Search for: Google Calendar API
* Select Google Calendar API.
* Click Enable.
3. Configure OAuth Authentication
* Open: Google Auth Platform -> https://console.cloud.google.com/auth/branding?project=calendar-api-510113
* Complete the required Branding information:
	Application Name
	Support Email
	Developer Contact Email
* Save the configuration.
4. Configure User Access
* Open: Google Auth Platform → Audience
* Select: Testing
* Add the Google account(s) that will use the tool as: Test Users
  For personal use, simply add your own Google account.

5. Create OAuth Credentials
* Open: Google Auth Platform → Clients
* Click: Create Client
* Choose: Desktop Application
* Enter a name such as: Outlook2Gcal Desktop Client
* Click Create.

Important: The client type must be Desktop Application. Do not choose Web Application.

6. Download the Credentials File
* Download the generated OAuth client JSON file.
* Rename the file: credentials.json
* Save it in the same folder as: outlook2gcal_api.py

Note:
If the app is in Testing Publishing status, then the token will need to be reauthorized for the Google Calendar API approximately every 7 days.

To reauthroize the token.json, run the following commands in whatever directory the app is stored.
* Remove-Item token.json
* py outlook2gcal_api.py

To change the app to In production so that the app should not have to be reauthorized, then on the Google Author Platform->Branding page:
* Complete the fields under App domain.
* enter Authorized domain, for example, github.com.
* Go back to the Audience page and the In production option should now be available. Click it!

One possible method for these additional fields on the Branding page is to create a GitHub repo like this one with Privacy.md and Terms.md files. Copy the links from the repo and files into the appropriate fields on the Branding page.

7. Use Windows Task Scheduler to run automatically - https://github.com/hotlanta/Outlook2Gcal/blob/main/documents/task_scheduler.md
