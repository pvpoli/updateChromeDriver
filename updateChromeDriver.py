from win32api import GetFileVersionInfo, HIWORD
import requests
import os
import zipfile
from pathlib import Path

class UpdateChromeDriver:

    DRIVER_FILES = ['chromedriver.exe', 'LICENSE.chromedriver', 'THIRD_PARTY_NOTICES.chromedriver']
    SUPPORTED_PLATFORMS = ["linux64", "mac-arm64", "mac-x64", "win32","win64"]

    def update_chrome_driver(self,destination, platform = "win64"):
        """
        Method to download and extract the last stable version of google chrome.
        It will first check the version of the chromedriver.exe file on the destination folder, if any.
        The download won't take place if the version of the chromedriver.exe file is the same as the last
        available stable version.

        ! ! ! Attention ! ! ! 
        It will delete any older versions of chromedriver from de destination folder!

        Arguments:
        destination: the directory on which to extract the chromedriver files
        platform: one of the supported platforms
        """
        if platform not in self.SUPPORTED_PLATFORMS:
            return (f"Aborting. The desired platform {platform} is not supported. Supported platforms: {self.SUPPORTED_PLATFORMS}")
        
        current_version = self.get_version(f"{destination}/chromedriver.exe")
        availlable = self.get_last_stable_download_link_from_data(platform)
        availlable_version = availlable.get("version")
        if current_version!=availlable_version:
            self.delete_old(destination)
            print(f"Updating chrome driver from version {current_version} to {availlable["version"]}")
            self.download_file(availlable["URL"],destination)
            self.extract(destination,platform)
            self.delete_zip(destination)
        else:
            print(f"You chrome driver current version {current_version} is the same as the last stable version: {availlable_version}. Nothing to update.")

    def get_version(self,destination):
        if os.path.isfile(destination):
            fileVersion = GetFileVersionInfo(destination,"\\")
            version = str(HIWORD(fileVersion["FileVersionMS"]))
            return version
        else:
            return "None"

    def get_data(self):
        url = "https://googlechromelabs.github.io/chrome-for-testing/last-known-good-versions-with-downloads.json"
        response = requests.get(url)
        return response.json()

    def get_last_stable_download_link_from_data(self, platform = "win64"):
        data = self.get_data()
        version = data.get("channels", {}).get("Stable",{}).get("version")
        version = version.split(".")[0]
        link = [chrome["url"] for chrome in data.get("channels", {}).get("Stable",{}).get("downloads", {}).get("chromedriver") if chrome["platform"]==platform]
        return {"version":version, "URL":link[0]}

    def download_file(self, downloadURL, destination, platform = "win64", chunk_size=128):
        response = requests.get(downloadURL, stream=True)
        with open(f"{destination}/chromedriver-{platform}.zip", 'wb') as fd:
            for chunk in response.iter_content(chunk_size=chunk_size):
                fd.write(chunk)
    
    def extract(self, destination, platform = "win64"):
        with zipfile.ZipFile(f"{destination}/chromedriver-{platform}.zip", 'r') as zip_ref:
            zip_ref.extractall(destination)
        tmp_folder = f"{destination}/chromedriver-{platform}"
        files = [f for f in os.listdir(tmp_folder) if os.path.isfile(os.path.join(tmp_folder, f))]
        for file in files:
            tmp_filePath = f"{tmp_folder}/{file}"
            if os.path.isfile(tmp_filePath):
                filePath = f"{destination}/{file}"
                os.rename(tmp_filePath,filePath)

    def delete_old(self, destination):
        print("Cleaning old files")
        files = [f for f in os.listdir(destination) if os.path.isfile(os.path.join(destination, f))]
        for file in files:
            if file in self.DRIVER_FILES:
                filePath = f"{destination}/{file}"
                if os.path.isfile(filePath):
                    print(f"Deleting file {file} from {destination}")
                    Path.unlink(filePath)
    
    def delete_zip(self,destination, platform = "win64"):
        print(f"Cleaning folder {destination}")
        Path.unlink(f"{destination}/chromedriver-{platform}.zip")
        os.rmdir(f"{destination}/chromedriver-{platform}")

if __name__ == "__main__":
    destination = input("Enter folder destination: ") or os.getcwd()
    print(f"Using {destination} as install destination.")
    patform = input("Enter the platform for which you want chromedriver (win32, win64, linux64, mac-arm64, mac-x64): ") or "win64"
    print(f"Updating chromedriver for platform: {patform}")
    UpdateChromeDriver().update_chrome_driver(destination, platform = "win64")