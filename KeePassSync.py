# coding=utf-8

import dropbox
import os.path
import subprocess
import os
import re
import requests
import win32api
import tempfile
import zipfile
import time
import argparse

def get_file_win_properties(fname):
    propNames = ('Comments', 'InternalName', 'ProductName',
        'CompanyName', 'LegalCopyright', 'ProductVersion',
        'FileDescription', 'LegalTrademarks', 'PrivateBuild',
        'FileVersion', 'OriginalFilename', 'SpecialBuild')

    props = {'FileVersion': "0.0.0"}

    try:
        # backslash as parm returns dictionary of numeric info corresponding to VS_FIXEDFILEINFO struc
        fixedInfo = win32api.GetFileVersionInfo(fname, '\\')
        props['FixedFileInfo'] = fixedInfo
        props['FileVersion'] = "%d.%d.%d.%d" % (fixedInfo['FileVersionMS'] / 65536,
                fixedInfo['FileVersionMS'] % 65536, fixedInfo['FileVersionLS'] / 65536,
                fixedInfo['FileVersionLS'] % 65536)

        # \VarFileInfo\Translation returns list of available (language, codepage)
        # pairs that can be used to retreive string info. We are using only the first pair.
        lang, codepage = win32api.GetFileVersionInfo(fname, '\\VarFileInfo\\Translation')[0]

        strInfo = {}
        for propName in propNames:
            strInfoPath = u'\\StringFileInfo\\%04X%04X\\%s' % (lang, codepage, propName)
            strInfo[propName] = win32api.GetFileVersionInfo(fname, strInfoPath)

        props['StringFileInfo'] = strInfo
    except:
        pass

    return props

def remove_trailing_zeroes(s):
    return re.sub(r"\.0\.0$", "", s)

def get_version(path):
    version = get_file_win_properties(path)['FileVersion']
    version = remove_trailing_zeroes(version)
    return version

def get_server_version(major_version):
    print("Checking for updates...")
    if major_version == 1:
        response = requests.Session().get("https://keepass.info/update/version1x.txt")
        #KeePass#1.38.0.0
        r = re.search(r"KeePass#([^\n^\r]*)", response.text)
        version = r.group(1)
        version = remove_trailing_zeroes(version)
    else:
        response = requests.Session().get("https://keepass.info/update/version2x.txt")
        #KeePass:2.46
        r = re.search(r"KeePass:([^\n^\r]*)", response.text)
        version = r.group(1)
    return version

def download_to_tmp(url):
    print("Downloading", url)
    session = requests.Session()
    response = session.head(url)
    response.raise_for_status()
    artifact_size = int(response.headers['Content-Length'])
    factor = 10
    chunk_size = int(artifact_size / factor)
    progress = 0
    with session.get(url, stream=True) as response:
        with tempfile.NamedTemporaryFile(delete=False) as f:
            for chunk in response.iter_content(chunk_size=chunk_size):
                if chunk:
                    print('Downloading: {}%'.format(progress))
                    progress += factor
                    f.write(chunk)
            return f.name
    return None

def extract_zip(path_to_zip, out_dir):
    with zipfile.ZipFile(path_to_zip) as z:
        for f in z.infolist():
            if f.is_dir():
                continue
            name, date_time = f.filename, f.date_time
            name = os.path.join(out_dir, os.path.normpath(name))
            if not os.path.exists(os.path.dirname(name)):
                os.makedirs(os.path.dirname(name))
            with open(name, 'wb') as out_file:
                out_file.write(z.open(f).read())
            date_time = time.mktime(date_time + (0, 0, -1))
            os.utime(name, (date_time, date_time))

def update_exe(folder):
    file_version = get_version(os.path.join(folder, "KeePass.exe"))
    major_version = int(file_version[0])
    server_version = get_server_version(major_version)
    if file_version == server_version:
        print("Exe is up to date")
        return
    
    print("Updating: %s -> %s" % (file_version, server_version))
    version = server_version
    while version[-2:] == ".0":
        version = version[:-2]
    download_lnk = "https://netix.dl.sourceforge.net/project/keepass/KeePass%%20%(major_version)d.x/%(version)s/KeePass-%(version)s.zip" % {"version": version, "major_version": major_version}
    path_to_zip = download_to_tmp(download_lnk)
    extract_zip(path_to_zip, folder)
    os.remove(path_to_zip)
    print("Done")

def dropbox_file_exists(dbx, path):
    try:
        dbx.files_get_metadata(path)
        return True
    except:
        return False

def main(args):
    update_exe(args.folder)

    dbx = dropbox.Dropbox(args.token)
    dbx.users_get_current_account()

    path_to_kdb_local = os.path.abspath(args.kdb_path)
    path_to_kdb_dropbox = args.dropbox_folder + os.path.basename(path_to_kdb_local)
    kdb_exists_in_dropbox = dropbox_file_exists(dbx, path_to_kdb_dropbox)

    if kdb_exists_in_dropbox:
        print("Downloading kdb...",)
        dbx.files_download_to_file(path_to_kdb_local, path_to_kdb_dropbox)
        print("Done")
    else:
        print("Database doesn't exist in Dropbox")

    local_modification_time = None
    if os.path.exists(path_to_kdb_local):
        local_modification_time = os.path.getmtime(path_to_kdb_local)

    print("Open KeePass")
    p = subprocess.Popen((os.path.join(args.folder, "KeePass.exe"), path_to_kdb_local))
    print("Waiting for program to close...",)
    p.wait()
    print("Done")

    kdb_changed = local_modification_time != os.path.getmtime(path_to_kdb_local)

    if kdb_changed or not kdb_exists_in_dropbox:
        print("Kdb changed, uploading...")
        with open(path_to_kdb_local, 'rb') as f:
            file_contents = f.read()
        dbx.files_upload(file_contents, path_to_kdb_dropbox, mode=dropbox.files.WriteMode('overwrite', None))
        print("Done")
    else:
        print("Kdb not changed")

    if args.remove_local_kdb:
        print("Removing kdb...",)
        os.remove(path_to_kdb_local)
        print("Done")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("-t", "--token", help="dropbox api token", required=True)
    parser.add_argument("-f", "--folder", help="path to keepass folder", default=".")
    parser.add_argument("-kdb", "--kdb-path", help="path to keepass kdb", default="Database.kdbx")
    parser.add_argument("-df", "--dropbox-folder", help="dropbox folder to database file", default="/")
    parser.add_argument("-r", "--remove-local-kdb", action="store_true", default=False, help="remove kdb from local folder after syncing")

    args = parser.parse_args()

    main(args)
