Bring Your Own Credentials
=========================================

In order to start using ``gslides`` you need to supply the package with Google API credentials. To set up a connection to a Google API, one must register an application with Google and make a decision whether they wish to utilize API keys, OAuth 2.0 client ID's or service accounts.

Below will walk you through how to set up a connection using OAuth 2.0 on local in a Jupyter Notebook.

Running on local with a Jupyter notebook
-----------------------------------------

This documentation follow roughly the steps laid out by Google `here <https://developers.google.com/slides/quickstart/python>`_.

**1. Create a project and enable the APIs**

You'll need enable a google developer account to access the Google Cloud Platform. From there you can follow this `guide <https://developers.google.com/workspace/guides/create-project>`_ to create a project and enable an API.

.. note::
   You will need to create a project with the Google Slides and Sheets API enabled

**2. Create credentials**

Again consult Google and follow this `guide <https://developers.google.com/workspace/guides/create-credentials#configure_the_oauth_consent_screen>`_ to configure the OAuth consent screen.

And create your credentials from the instructions `here <https://developers.google.com/workspace/guides/create-credentials#create_a_credential>`_. You'll need to download the credentials ``.json`` file onto your local. *Be sure not share these credentials with anyone, the file should not be shared on Github*.

**3. Update the URIs**

The URI are the endpoints where the OAuth 2.0 server can send responses. In this case because we are utilizing a jupyter notebook we must allow responses to ``localhost``. Update the values like so, selecting a port which you are going to use.

.. image:: img/uri.png

**4. Launch a Jupyter notebook, get the token and load credentials**

Run the code below, replacing the necessary values to obtain the ``creds`` object. These are the credentials you pass to ``gslides``.

.. code-block:: python

  import os.path
  from googleapiclient.discovery import build
  from google_auth_oauthlib.flow import InstalledAppFlow
  from google.auth.transport.requests import Request
  from google.oauth2.credentials import Credentials

  # These scopes are read & write permissions. Necessary to run gslides
  SCOPES = ['https://www.googleapis.com/auth/presentations',
           'https://www.googleapis.com/auth/spreadsheets']

  creds = None
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  if os.path.exists('token.json'):
      creds = Credentials.from_authorized_user_file('token.json', SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
      if creds and creds.expired and creds.refresh_token:
          creds.refresh(Request())
      else:
          flow = InstalledAppFlow.from_client_secrets_file(
              '<PATH_TO_CREDS>', SCOPES)
          creds = flow.run_local_server(port=<PORT>)
      # Save the credentials for the next run
      with open('token.json', 'w') as token:
          token.write(creds.to_json())

**5. Initialize the credentials with gslides**

.. code-block:: python

  gslides.intialize_credentials(creds)