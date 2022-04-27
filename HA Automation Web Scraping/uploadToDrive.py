from pydrive.drive import GoogleDrive
from pydrive.auth import GoogleAuth
import os

gauth = GoogleAuth()
gauth.LocalWebserverAuth()
current_path = os.getcwd()
print(current_path)
drive = GoogleDrive(gauth)
filename = 'Artists.csv'
f = drive.CreateFile({'title': filename})

f.SetContentFile(os.path.join(current_path, filename))
f.Upload()