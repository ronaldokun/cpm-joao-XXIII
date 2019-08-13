# ---
# jupyter:
#   jupytext:
#     text_representation:
#       extension: .py
#       format_name: light
#       format_version: '1.3'
#       jupytext_version: 0.8.6
#   kernelspec:
#     display_name: Python 3
#     language: python
#     name: python3
# ---

from __future__ import print_function
import httplib2
import os, io

import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
import auth

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/drive-python-quickstart.json
SCOPES = "https://www.googleapis.com/auth/drive"
CLIENT_SECRET_FILE = "drive_joao23.json"
APPLICATION_NAME = "Drive API Python"
authInst = auth.auth(SCOPES, CLIENT_SECRET_FILE, APPLICATION_NAME)
drive_service = authInst.getCredentials()

# http = credentials.authorize(httplib2.Http())
# drive_service = build('drive', 'v3', http=credentials)


def listFiles(size):
    results = (
        drive_service.files()
        .list(pageSize=size, fields="nextPageToken, files(id, name)")
        .execute()
    )
    items = results.get("files", [])
    if not items:
        print("No files found.")
    else:
        print("Files:")
    for item in items:
        print("{0} ({1})".format(item["name"], item["id"]))


def uploadFile(filename, filepath, mimetype):
    file_metadata = {"name": filename}
    media = MediaFileUpload(filepath, mimetype=mimetype)
    file = (
        drive_service.files()
        .create(body=file_metadata, media_body=media, fields="id")
        .execute()
    )
    print("File ID: %s" % file.get("id"))


def downloadFile(file_id, filepath):
    request = drive_service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print("Download %d%%." % int(status.progress() * 100))
    with io.open(filepath, "wb") as f:
        fh.seek(0)
        f.write(fh.read())


def createFolder(name):
    file_metadata = {"name": name, "mimeType": "application/vnd.google-apps.folder"}
    file = drive_service.files().create(body=file_metadata, fields="id").execute()
    print("Folder ID: %s" % file.get("id"))


def searchFile(size, query):
    results = (
        drive_service.files()
        .list(
            pageSize=size,
            fields="nextPageToken, files(id, name, kind, mimeType)",
            q=query,
        )
        .execute()
    )
    items = results.get("files", [])
    if not items:
        print("No files found.")
    else:
        print(len(items))
        return items


# uploadFile('unnamed.jpg','unnamed.jpg','image/jpeg')
# downloadFile('1Knxs5kRAMnoH5fivGeNsdrj_SIgLiqzV','google.jpg')
# createFolder('Google')
searchFile(100, "mimeType contains 'image'")

