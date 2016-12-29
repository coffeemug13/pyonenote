# pyonenote
Experimental python wrapper for OneNote. This was a quick hack to migrate data from my Journey Two App to Microsoft OneNote. Because I wanted to use python but I didn't found a working lib, which I could use at this date just as command line tool, I wrote my own. By the way getting a better understanding of OAuth2 and MS Graph.

Unfortunately the MS Graph Beta API for OneNote is not yet working, therefore I had to fall back to the old API.

A good place to start reading is:
https://msdn.microsoft.com/en-us/office/office365/howto/onenote-auth#onenote-perms-msa

Play around with the API
http://dev.onenote.com/docs#/reference/get-notebooks/v10menotesnotebooksidselectexpand/get?console=1

# Setup
- register your App at https://account.live.com/developers/applications and get your *client_id and client_secret*
- open the URL from OneNote.get_authorize_url() in a browser to grant access of this app to the user notebooks
- extract (manually) from the redirect URL the authentication *code*
- get your initial token (not yet implemented but with comments in OneNote.get_token())
- make your requests
