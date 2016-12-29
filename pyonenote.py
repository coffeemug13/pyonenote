"""
PyOneNote.py
~~~~~~~~~~~~~~~~~

This module contains a basic OAuth 2 Authentication and basic handler for GET and POST operations.
This work was just a quick hack to migrate notes from and old database to onenote but should hep you to understand
the request structure of OneNote.

Copyright (c) 2016 Coffeemug13. All rights reserved. Licensed under the MIT license.
See LICENSE in the project root for license information.
"""

import requests


class OAuth():
    """Handles the authentication for all requests"""

    def __init__(self, client_id, client_secret, code=None, token=None, refresh_token=None):
        """ This information is obtained upon registration of a new Outlook Application
         The values are just for information and not valid
         :param client_id: "cda3ffaa-2345-a122-3454-adadc556e7bf"
         :param client_secret: "AABfsafd6Q5d1VZmJQNsdac"
         :param code: = "AcD5bcf9a-0fef-ca3a-1a3a-9v4543388572"
         :param token: = "EAFSDTBRB$/UGCCXc8wU/zFu9QnLdZXy+YnElFkAAW......"
         :param rtoken: = "MCKKgf55PCiM2aACbIYads*sdsa%*PWYNj436348v......" """
        self.client_id = client_id
        self.client_secret = client_secret
        self.code = code
        self.token = token
        self.rtoken = refresh_token
        self.redirect_uri = 'https://localhost'
        self.session = requests.Session()

    @staticmethod
    def get_authorize_url(client_id):
        "open this url in a browser to let the user grant access to onenote. Extract from the return URL your access code"
        url = "https://login.live.com/oauth20_authorize.srf?client_id={0}&scope=wl.signin%20wl.offline_access%20wl.basic%20office.onenote_create&response_type=code&redirect_uri=https://localhost".format(
            client_id)
        return url

    def get_token(self):
        """
        Make the following request with e.g. postman:
        POST https://login.live.com/oauth20_token.srf
        Content-Type:application/x-www-form-urlencoded

        grant_type:authorization_code
        client_id:cda3ffaa-2345-a122-3454-adadc556e7bf
        client_secret:AABfsafd6Q5d1VZmJQNsdac
        code:111111111-1111-1111-1111-111111111111
        redirect_uri:https://localhost
        
        OneNote will return as result:
        {
          "token_type": "bearer",
          "expires_in": 3600,
          "scope": "wl.signin wl.offline_access wl.basic office.onenote_create office.onenote",
          "access_token": "AxxdWR1DBAAUGCCXc8wU/....",
          "refresh_token": "DR3DDEQJPCiM2aACbIYa....",
          "user_id": "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
        }
        """
        raise NotImplementedError("")

    def refresh_token(self):
        """
        Make the following reqest to refresh you token with e.g. postman:
        POST https://login.live.com/oauth20_token.srf
        Content-Type:application/x-www-form-urlencoded

        grant_type:refresh_token
        client_id:cda3ffaa-2345-a122-3454-adadc556e7bf
        client_secret:AABfsafd6Q5d1VZmJQNsdac
        refresh_token:DR3DDEQJPCiM2aACbIYa....
        redirect_uri:https://localhost
        -->
        {
          "token_type": "bearer",
          "expires_in": 3600,
          "scope": "wl.signin wl.offline_access wl.basic office.onenote_create office.onenote",
          "access_token": "EAFSDTBRB$/UGCCXc8wU/zFu9QnLdZXy+YnElFkAAW...",
          "refresh_token": "DSFDSGSGFABDBGFGBFGF5435kFGDd2J6Bco2Pv2ss...",
          "user_id": "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
        }
        """
        url = 'https://login.live.com/oauth20_token.srf'
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        data = {"grant_type": "refresh_token",
                "client_id": self.client_id,
                "client_secret": self.client_secret,
                "refresh_token": self.rtoken,
                "redirect_uri": self.redirect_uri}

        result = self.session.post(url, headers=headers, data=data)

        print("Refreshed token: " + result.text)
        refresh = result.json()
        self.expire = refresh.get('expires_in')
        self.token = refresh.get('access_token')
        self.rtoken = refresh.get('refresh_token')
        print("Token: " + self.token)
        print("Refresh Token: " + self.rtoken)
        return True

    def _get(self, url, query):
        """Handles GET Request with Authentication"""
        headers = {'user-agent': 'my-app/0.0.1', 'Authorization': 'Bearer ' + self.token}
        result = self.session.get(url, headers=headers, params=query)
        print("GET " + result.url)
        print(result.headers)
        if (result.text):
            print(result.text)
        return result

    def _post(self, url: str, headers: list, data: str = None, files: list = None):
        """Handles POST Request with Authentication"""
        newHeaders = {'user-agent': 'my-app/0.0.1', 'Authorization': 'Bearer ' + self.token}
        if data:
            newHeaders.update(headers)
            result = self.session.post(url, headers=newHeaders, data=data)
        else:
            result = self.session.post(url, headers=newHeaders, files=files)
        # result.request.headers
        print("POST " + result.url)
        print(result.headers)
        if (result.text):
            print(result.text)
        return result

    def post(self, url: str, headers: list, data: str = None, files: list = None):
        """post something and handle token expire transparent to the caller"""
        try:
            result = self._post(url, headers, data=data, files=files)
            if (result.status_code not in (200, 201)):
                print("Error: " + str(result.status_code))
                if (result.status_code == 401):
                    print("Refreshing token")
                    if self.refresh_token():
                        result = self._post(url, headers, data, files=files)
                    else:
                        print('Failed retry refreshing token')
            return result
        except Exception as e:
            print(e)
            pass

    def get(self, url, query, headers=None):
        """get something and handle token expire transparent to the caller"""
        try:
            result = self._get(url, query)
            if (result.status_code != requests.codes.ok):
                print("Error: " + str(result.status_code))
                if (result.status_code == 401):
                    print("Refreshing token")
                    if self.refresh_token():
                        result = self._get(url, query)
                    else:
                        print('Failed retry refreshing token')
            return result
        except Exception as e:
            print(e)
            pass

    def get_credentials(self):
        """Return the actual credentials of this OAuth Instance
        :return client_id:"""
        return self.client_id, self.client_secret, self.code, self.token, self.rtoken


class OneNote(OAuth):
    """This class wraps some OneNote specific calls"""
    def __init__(self, client_id, client_secret, code, token, rtoken):
        super().__init__(client_id, client_secret, code, token, rtoken)
        self.base = "https://www.onenote.com/api/v1.0/me/"

    def list_notebooks(self):
        url = self.base + "notes/notebooks"
        query = {'top': '5'}
        result = self.get(url, query)
        n = None
        if (result):
            notebooks = result.json()
            # result_serialized = json.dumps(result.text)
            # notebook = json.loads(result_serialized)
            n = notebooks["value"][0]
            x = 1
        return n

    def post_page(self, section_id: str, created, title: str, content: str, files: list = None):
        """post a page. If you want to provide additional images to the page provide them as file list
        in the same way like posting multipart message in 'requests'
        .:param content: valid html text with Umlaute converted to &auml;"""
        url = self.base + "notes/sections/" + section_id + "/pages"
        headers = {"Content-Type": "application/xhtml+xml"}
        # the basic layout of a page is always same
        data = """<?xml version="1.0" encoding="utf-8" ?>
<html>
  <head>
    <title>{0}</title>
    <meta name="created" content="{1}"/>
  </head>
  <body data-absolute-enabled="true">
    <div>
      {2}
    </div>
  </body>
</html>
""".format(title, created, content)
        result = None
        if files:
            "post as multipart"
            newFiles = [('Presentation', (None, data, 'application/xhtml+xml', {'Content-Encoding': 'utf8'}))]
            newFiles.extend(files)
            result = self.post(url, {}, None, files=newFiles)
        else:
            "post as simple request"
            result = self.post(url, headers, data)
        n = None
        if (result):
            notebooks = result.json()
            # result_serialized = json.dumps(result.text)
            # notebook = json.loads(result_serialized)
            # n = notebooks["value"][0]
            x = 1
        return notebooks
