Bring Your Own Credentials
=========================================

In order to start using ``gslides`` you need to supply the package with Google API credentials. To set up a connection to a Google API, one must register an application with Google and make a decision whether they wish to utilize API keys, OAuth 2.0 client ID's or service accounts.

Below will walk you through how to set up a connection using OAuth 2.0 or a service account on local in a Jupyter Notebook.

Running on local with a Jupyter notebook
-----------------------------------------

This documentation follow roughly the steps laid out by Google `here <https://developers.google.com/slides/quickstart/python>`_.

**1. Create a project and enable the APIs**

You'll need enable a google developer account to access the Google Cloud Platform. From there you can follow this `guide <https://developers.google.com/workspace/guides/create-project>`_ to create a project and enable an API.

.. note::
   You will need to create a project with the Google Slides and Sheets API enabled

**2. Create credentials**

You can either create credentials with OAuth or a service account. See commentary in :ref:`Service account or OAuth 2.0?`

*2a. Using OAuth*

Consult Google and follow this `guide <https://developers.google.com/workspace/guides/create-credentials#configure_the_oauth_consent_screen>`_ to configure the OAuth consent screen. This is the screen that will appear when you attempt to use your credentials to retrieve a token.

Next, create your credentials from the instructions `here <https://developers.google.com/workspace/guides/create-credentials#create_a_credential>`_. For simplicity, create a ``Desktop app``. Once finished, you'll need to download the credentials ``.json`` file onto your local.

.. warning::
  Be sure not share your credentials with anyone. The file should not be pushed to Github.

*2b. Using a service account*

To create a service account in your project navigate to ``--> Credentials --> Create Credentials --> Service Account``. Once you create the service account click into the account navigate to ``--> Keys --> Add Key``. You will have automatically downloaded your new key onto local.

**3. Launch a Jupyter notebook, get the token and load credentials**

*3b. Using OAuth*

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
          creds = flow.run_local_server()
      # Save the credentials for the next run
      with open('token.json', 'w') as token:
          token.write(creds.to_json())

The first time you run this block, you will be prompted to allow access to Google slides & sheets through your Google account.

*3b. Using a service account*

.. code-block:: python

  from google.oauth2 import service_account

  SCOPES = ['https://www.googleapis.com/auth/presentations',
           'https://www.googleapis.com/auth/spreadsheets']

  credentials = service_account.Credentials.from_service_account_file(
      '<PATH_TO_CREDS>')

  creds = credentials.with_scopes(SCOPES)

**4. Initialize the credentials with gslides**

.. code-block:: python

  gslides.initialize_credentials(creds)


Service account or OAuth 2.0?
-----------------------------------------

The key difference between the usage of a service account or OAuth 2.0 to use the google API is the google account that making the API requests. Using OAuth 2.0 **the account you authorize as will be the account that makes the API requests**. This means that any presentations or spreadsheets you create will show up in your own Google drive. If you leverage a service account the presentation or spreadsheet will be created by the service account, an account with the domain like so ``@<project>.iam.gserviceaccount.com``. If you try to navigate to the presentation under your own gmail account you will not have access, this is because the service account is owner for that document. This doesn't prevent you from leveraging ``gslides`` with a service account although. For any existing documents you would like to edit with ``gslides`` simply share the document with the service account email, setting the account as editor.
