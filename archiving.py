import os
import zipfile
import fnmatch

class ArchivingApp:
    def __init__(self,
                compress_source_folder=None,
                compress_target_file=None,
                uncompress_source_file=None,
                uncompress_target_folder=None,
                file_types=None):
        self.target_file = compress_target_file
        self.target_zip = zipfile.ZipFile(self.target_file, 'w')
        self.compress_source = compress_source_folder
        self.source_zip = uncompress_source_file
        self.target_folder = uncompress_target_folder
        self.types = file_types

    def compress(self,protect_key):
        try:
            for folder, subfolders, files in os.walk(self.compress_source):
                for file in files:
                    for pattern in self.types:
                        if fnmatch.fnmatch(file,pattern):
                            self.target_zip.write(os.path.join(folder, file),
                            os.path.relpath(os.path.join(folder,file), self.compress_source), compress_type = zipfile.ZIP_DEFLATED)
            if protect_key != None:
                self.target_zip.setpassword(protect_key)
        except Exception as e:
            return f"Error while archiving - {e}"

        self.target_zip.close()
        return f"Matching files have successfully been compressed into {self.target_file}"

    def remove(self):
        cd = os.getcwd()
        os.chdir(self.compress_source)
        for folder, subfolders, files in os.walk(self.compress_source):
            for file in files:
                for pattern in self.types:
                    if fnmatch.fnmatch(file,pattern):
                        os.remove(file)
        os.chdir(cd)
        return f"removed"

    def uncompress(self):
        try:
            source_zip = zipfile.ZipFile(self.source_zip)
            source_zip.extractall(self.target_folder)
            return f"{self.source_zip} has successfully been decompressed into {self.target_folder}"
        except Exception as e:
            return f"Eorror: {e}"

def main():
    source_fld = "c:\\MindTech\\KWP"
    target = "mindtechs.zip"
    types = ['*.csv','B*.xlsx']
    app = ArchivingApp(compress_source_folder=source_fld,compress_target_file=target,file_types=types)
    app.compress("kantar@123")

if __name__ == "__main__":
    main()
