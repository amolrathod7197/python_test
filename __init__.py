import sys
import win32com.client
import win32file
import magic
import os
from datetime import datetime
import array
from ctypes import *
import re



sys.path.insert(0, os.getcwd())
from helper import Logger

OPTION_VER_NO = "--ver"
OPTION_CHECK_AGENT = "--agent"
OPTION_DIST_PATH = "--path"
LOG_FILENAME = "AgentBuildValidator.log"
sys.stdout = Logger(LOG_FILENAME)

class Agent_Build_Validator:
    def __init__(self,path,ver_no,agent):
        print "Inside init"
        self.fail = False
        self.path= path
        self.agent= agent
        self.ver_no = ver_no

        self.Vaultize_Agent_func_list = [self.check_Vaultize_agent,
                                        self.check_digital_signature_Vaultize, self.check_file_type_and_bitness_Vaultize,
                                        self.validate_build_number_Vaultize, self.check_sizeof_Vaultize_Agent,
                                        self.check_Timestamp_Vaultize]

        self.Outlook_Addin_func_list =  [self.check_Vaultize_Outlook_Addin,
                                        self.check_digital_signature_Outlook_addin, self.check_Timestamp_Outlook_Addin,
                                        self.check_sizeof_Outlook_Addin, self.check_file_type_and_bitness_Outlook_Addin,
                                        self.validate_build_number_Outlook_Addin]

        self.vDRM_func_list =   [self.check_vDRM_Agent, self.check_file_type_and_bitness_vDRM,
                                self.check_digital_signature_vDRM, self.check_sizeof_vDRM,
                                self.check_Timestamp_vDRM, self.validate_build_number_vDRM]

        self.ALL_func_list = [self.check_Vaultize_agent, self.check_digital_signature_Vaultize,
                                self.check_file_type_and_bitness_Vaultize, self.validate_build_number_Vaultize,
                                self.check_sizeof_Vaultize_Agent, self.check_Timestamp_Vaultize,
                                self.check_Vaultize_Outlook_Addin, self.check_digital_signature_Outlook_addin,
                                self.check_sizeof_Outlook_Addin, self.check_file_type_and_bitness_Outlook_Addin,
                                self.check_Timestamp_Outlook_Addin, self.validate_build_number_Outlook_Addin,
                                self.check_vDRM_Agent, self.check_file_type_and_bitness_vDRM,
                                self.check_digital_signature_vDRM,self.check_sizeof_vDRM,
                                self.validate_build_number_vDRM, self.check_Timestamp_vDRM,
                                self.check_Timestamp_AD_Notes, self.validate_build_number_AD_Notes,
                                self.check_file_type_and_bitness_AD_Notes,self.check_sizeof_AD_Notes,
                                self.check_AD_Notes_plugin,self.check_digital_signature_AD_Notes_plugin,
                                self.check_file_name_pattern]

    def run(self):
        print "Running validation"
        try:
            if self.agent=="Vaultize_Agent":
                for f in self.Vaultize_Agent_func_list:
                    f()
                    if self.fail == True:
                        print "*** FAIL detected ***"
                        break
            elif self.agent=="Outlook_Addin":
                for f in self.Outlook_Addin_func_list:
                    f()
                    if self.fail == True:
                        print "*** FAIL detected ***"
                        break
            elif self.agent=="vDRM":
                for f in self.vDRM_func_list:
                    f()
                    if self.fail == True:
                        print "*** FAIL detected ***"
                        break
            elif self.agent=="ALL":
                for f in self.ALL_func_list:
                    f()
                    if self.fail == True:
                        print "*** FAIL detected ***"
                        break

        finally:
            self.cleanup()
        return self.fail #This will translate to exit code


    def check_Vaultize_agent(self):#check for file
        try:
            print("checking for vaultize agent...")
            Vaultize_setup_exe = os.path.join(self.path, 'Vaultize-' + self.ver_no + '-Setup.exe')
            if os.path.exists(Vaultize_setup_exe):
                if os.path.isfile(Vaultize_setup_exe):
                    pass
                else:
                    print(Vaultize_setup_exe+" is not a File.")
                    self.fail=True
                print('Vaultize-' + self.ver_no + '-Setup.exe is present')

            else:
                print('Vaultize-' + self.ver_no + '-Setup.exe is not present')

            Vaultize_Silent_setup = os.path.join(self.path, 'Vaultize-' + self.ver_no + '-Silent-Setup.exe')
            if os.path.exists(Vaultize_Silent_setup):
                if os.path.isfile(Vaultize_Silent_setup):
                    pass
                else:
                    print(Vaultize_Silent_setup+" is not a File.")
                    self.fail=True
                print('Vaultize-' + self.ver_no + '-Silent-Setup.exe is present')
            else:
                print('Vaultize-' + self.ver_no + '-Silent-Setup.exe is not present')
                self.fail=True
        except Exception as e:
            print(e)
    def check_Vaultize_Outlook_Addin(self):
        print("checking for Vaultize Outlook Addin....")

        Outlook_addin_exe=os.path.join(self.path,'Vaultize-'+self.ver_no+'-Outlook-Addin-Installer.exe')
        if os.path.exists(Outlook_addin_exe):
            if os.path.isfile(Outlook_addin_exe):
                pass
            else:
                print(Outlook_addin_exe + " is not a File.")
                self.fail = True
            print("Vaultize-"+self.ver_no+"-Outlook-Addin-Installer.exe is present")
        else:
            print("Vaultize-"+self.ver_no+"-Outlook-Addin-Installer.exe is Not present.")
            self.fail=True

        Outlook_addin_msi=os.path.join(self.path,'Vaultize-'+self.ver_no+'-Outlook-Addin-Installer.msi')
        if os.path.exists(Outlook_addin_msi):
            if os.path.isfile(Outlook_addin_msi):
                pass
            else:
                print(Outlook_addin_msi + " is not a File.")
                self.fail = True
            print("Vaultize-"+self.ver_no+"-Outlook-Addin-Installer.msi is present")
        else:
            print("Vaultize-"+self.ver_no+"-Outlook-Addin-Installer.msi is Not present.")
            self.fail = True

        Outlook_addin64_exe=os.path.join(self.path,'Vaultize-'+self.ver_no+'-Outlook-Addin-x64-Installer.exe')
        if os.path.exists(Outlook_addin64_exe):
            if os.path.isfile(Outlook_addin64_exe):
                pass
            else:
                print(Outlook_addin64_exe + " is not a File.")
                self.fail = True
            print("Vaultize-"+self.ver_no+"-Outlook-Addin-x64-Installer.exe is present")
        else:
            print("Vaultize-"+self.ver_no+"-Outlook-Addin-x64-Installer.exe is Not present.")
            self.fail = True
        Outlook_addin32_exe=os.path.join(self.path,'Vaultize-'+self.ver_no+'-Outlook-Addin-x86-Installer.exe')
        if os.path.exists(Outlook_addin32_exe):
            if os.path.isfile(Outlook_addin32_exe):
                pass
            else:
                print(Outlook_addin32_exe + " is not a File.")
                self.fail = True
            print("Vaultize-"+self.ver_no+"-Outlook-Addin-x86-Installer.exe is present")
        else:
            print("Vaultize-"+self.ver_no+"-Outlook-Addin-x86-Installer.exe is Not present.")
            self.fail = True

        Outlook_addin64_msi=os.path.join(self.path,'Vaultize-'+self.ver_no+'-Outlook-Addin-x64-Installer.msi')
        if os.path.exists(Outlook_addin64_msi):
            if os.path.isfile(Outlook_addin64_msi):
                pass
            else:
                print(Outlook_addin64_msi + " is not a File.")
                self.fail = True
            print("Vaultize-"+self.ver_no+"-Outlook-Addin-x64-Installer.msi is present")
        else:
            print("Vaultize-"+self.ver_no+"-Outlook-Addin-x64-Installer.msi is Not present.")
            self.fail = True

        Outlook_addin32_msi=os.path.join(self.path,'Vaultize-'+self.ver_no+'-Outlook-Addin-x86-Installer.msi')
        if os.path.exists(Outlook_addin32_msi):
            if os.path.isfile(Outlook_addin32_msi):
                pass
            else:
                print(Outlook_addin32_msi + " is not a File.")
                self.fail = True
            print("Vaultize-"+self.ver_no+"-Outlook-Addin-x86-Installer.msi is present")
        else:
            print("Vaultize-"+self.ver_no+"-Outlook-Addin-x86-Installer.msi is Not present.")
            self.fail = True

    def check_vDRM_Agent(self):
        print("checking for vDRM Agent...")
        vDRM_exe=os.path.join(self.path,'Vaultize-'+self.ver_no+'-vDRM-Setup.exe')
        if os.path.exists(vDRM_exe):
            if os.path.isfile(vDRM_exe):
                pass
            else:
                print(vDRM_exe + " is not a File.")
                self.fail = True
            print("Vaultize-"+self.ver_no+"-vDRM-Setup.exe is present")
        else:
            print("Vaultize-"+self.ver_no+"-vDRM-Setup.exe is Not present.")
            self.fail = True

        vDRM_msi=os.path.join(self.path,'Vaultize-'+self.ver_no+'-vDRM-Setup.msi')
        if os.path.exists(vDRM_msi):
            if os.path.isfile(vDRM_msi):
                pass
            else:
                print(vDRM_msi + " is not a File.")
                self.fail = True
            print("Vaultize-"+self.ver_no+"-vDRM-Setup.msi is present")
        else:
            print("Vaultize-"+self.ver_no+"-vDRM-Setup.msi is Not present.")
            self.fail = True

        vDRM_Standalone_exe=os.path.join(self.path,'Vaultize-'+self.ver_no+'-vDRM-Standalone.exe')
        if os.path.exists(vDRM_Standalone_exe):
            if os.path.isfile(vDRM_Standalone_exe):
                pass
            else:
                print(vDRM_Standalone_exe + " is not a File.")
                self.fail = True
            print("Vaultize-"+self.ver_no+"-vDRM-Standalone-Setup.exe is present")
        else:
            print("Vaultize-"+self.ver_no+"-vDRM-Standalone-Setup.exe is Not present.")
            self.fail = True

    def check_AD_Notes_plugin(self):
        Notes_plugin_setup = os.path.join(self.path,'Vaultize-' + self.ver_no + '-NotesPlugin-Setup.exe')
        if os.path.exists(Notes_plugin_setup):
            if os.path.isfile(Notes_plugin_setup):
                pass
            else:
                print(Notes_plugin_setup + " is not a File.")
                self.fail = True
            print("Vaultize-" + self.ver_no + "-NotesPlugin-Setup.exe present")
        else:
            print("Vaultize-" + self.ver_no + "-NotesPlugin-Setup.exe is Not present.")
            self.fail = True

        Vaultize_AD_setup_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-AD-Setup.exe')
        if os.path.exists(Vaultize_AD_setup_exe):
            if os.path.isfile(Vaultize_AD_setup_exe):
                pass
            else:
                print(Vaultize_AD_setup_exe + " is not a File.")
                self.fail = True
            print("Vaultize-" + self.ver_no + "-AD-Setup.exe present")
        else:
            print("Vaultize-" + self.ver_no + "-AD-Setup.exe is Not present.")
            self.fail = True

        Vaultize_AD_setup_msi = os.path.join(self.path,'Vaultize-' + self.ver_no + '-AD-Setup.msi')
        if os.path.exists(Vaultize_AD_setup_msi):
            if os.path.isfile(Vaultize_AD_setup_msi):
                pass
            else:
                print(Vaultize_AD_setup_msi + " is not a File.")
                self.fail = True
            print("Vaultize-" + self.ver_no + "-AD-Setup.msi present")
        else:
            print("Vaultize-" + self.ver_no + "-AD-Setup.msi is Not present.")
            self.fail = True

    def check_sizeof_Vaultize_Agent(self):
        print("checking size of Agents...")
        Vaultize_setup_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Setup.exe')
        b = os.path.getsize(Vaultize_setup_exe)
        b=b/(1024*1024)    #converting Byte to MegaByte

        if b in range(20,35):
            print('Vaultize-' + self.ver_no + '-Setup.exe size is ok')
        else:
            print('Vaultize-' + self.ver_no + '-Setup.exe is not in range..')
            self.fail = True

        Vaultize_silent = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Silent-Setup.exe')
        b = (os.path.getsize(Vaultize_silent))
        b=b/(1024*1024)
        if b in range(20,35):
            print('Vaultize-' + self.ver_no + '-Silent-Setup.exe size is ok')
        else:
            print('Vaultize-' + self.ver_no + '-Silent-Setup.exe size is not in range..')
            self.fail = True

    def check_sizeof_Outlook_Addin(self):
        Outlook_addin_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Outlook-Addin-Installer.exe')
        b = (os.path.getsize(Outlook_addin_exe))
        b=b/(1024*1024)
        if b in range(4,10):
            print('Vaultize-' + self.ver_no + '-Outlook-Addin-Installer.exe size is ok')
        else:
            print('Vaultize-' + self.ver_no + '-Outlook-Addin-Installer.exe size is not in range..')
            self.fail = True

        Outlook_addin64_msi = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Outlook-Addin-x64-Installer.msi')
        b = (os.path.getsize(Outlook_addin64_msi))
        b=b/(1024*1024)
        if b in range(2, 6):
            print('Vaultize-' + self.ver_no + '-Outlook-Addin-x64-Installer.msi size is ok')
        else:
            print('Vaultize-' + self.ver_no + '-Outlook-Addin-x64-Installer.msi size is not in range..')
            self.fail = True

        Outlook_addin32_msi = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Outlook-Addin-x86-Installer.msi')
        b = (os.path.getsize(Outlook_addin32_msi))
        b = b / (1024 * 1024)
        if b in range(2, 6):
            print('Vaultize-' + self.ver_no + '-Outlook-Addin-x86-Installer.msi size is ok')
        else:
            print('Vaultize-' + self.ver_no + '-Outlook-Addin-x86-Installer.msi size is not in range..')
            self.fail = True

        Outlook_addin32_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Outlook-Addin-x86-Installer.exe')
        b = (os.path.getsize(Outlook_addin32_exe))
        b = b / (1024 * 1024)
        if b in range(3, 6):
            print('Vaultize-' + self.ver_no + '-Outlook-Addin-x86-Installer.exe size is ok')
        else:
            print('Vaultize-' + self.ver_no + '-Outlook-Addin-x86-Installer.exe size is not in range..')
            self.fail = True

        Outlook_addin64_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Outlook-Addin-x64-Installer.exe')
        b = (os.path.getsize(Outlook_addin64_exe))
        b = b / (1024 * 1024)
        if b in range(3, 6):
            print('Vaultize-' + self.ver_no + '-Outlook-Addin-x64-Installer.exe size is ok')
        else:
            print('Vaultize-' + self.ver_no + '-Outlook-Addin-x64-Installer.exe size is not in range..')
            self.fail = True

        Outlook_addin_msi = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Outlook-Addin-Installer.msi')
        b = (os.path.getsize(Outlook_addin_msi))
        b = b / (1024 * 1024)
        if b in range(1, 4):
            print('Vaultize-' + self.ver_no + '-Outlook-Addin-Installer.msi size is ok')
        else:
            print('Vaultize-' + self.ver_no + '-Outlook-Addin-Installer.msi size is not in range..')
            self.fail = True

    def check_sizeof_vDRM(self):
        vDRM_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-vDRM-Setup.exe')
        b = (os.path.getsize(vDRM_exe))
        b = b / (1024 * 1024)
        if b in range(20, 30):
            print('Vaultize-' + self.ver_no + '-vDRM-Setup.exe size is ok')
        else:
            print('Vaultize-' + self.ver_no + '-vDRM-Setup.exe size is not in range..')
            self.fail = True

        vDRM_msi = os.path.join(self.path,'Vaultize-' + self.ver_no + '-vDRM-Setup.msi')
        b = (os.path.getsize(vDRM_msi))
        b = b / (1024 * 1024)
        if b in range(20, 30):
            print('Vaultize-' + self.ver_no + '-vDRM-Setup.msi size is ok')
        else:
            print('Vaultize-' + self.ver_no + '-vDRM-Setup.msi size is not in range..')
            self.fail = True

        vDRM_Standalone_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-vDRM-Standalone.exe')
        b = (os.path.getsize(vDRM_Standalone_exe))
        b = b / (1024 * 1024)
        if b in range(17, 25):
            print('Vaultize-' + self.ver_no + '-vDRM-Standalone.exe size is ok')
        else:
            print('Vaultize-' + self.ver_no + '-vDRM-Standalone.exe size is not in range..')
            self.fail = True

    def check_sizeof_AD_Notes(self):
        Notes_plugin_setup = os.path.join(self.path,'Vaultize-' + self.ver_no + '-NotesPlugin-Setup.exe')
        b = (os.path.getsize(Notes_plugin_setup))
        b = b / (1024 * 1024)
        if b in range(6, 12):
            print('Vaultize-' + self.ver_no + '-NotesPlugin-Setup.exe size is ok')
        else:
            print('Vaultize-' + self.ver_no + '-NotesPlugin-Setup.exe size is not in range..')
            self.fail = True

        Vaultize_AD_setup_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-AD-Setup.exe')
        b = (os.path.getsize(Vaultize_AD_setup_exe))
        b = b / (1024 * 1024)
        if b in range(0, 4):
            print('Vaultize-' + self.ver_no + '-AD-Setup.exe size is ok')
        else:
            print('Vaultize-' + self.ver_no + '-AD-Setup.exe size is not in range..')
            self.fail = True

        Vaultize_AD_setup_msi = os.path.join(self.path,'Vaultize-' + self.ver_no + '-AD-Setup.msi')
        b = (os.path.getsize(Vaultize_AD_setup_msi))
        b = b / (1024 * 1024)
        if b in range(34, 42):
            print('Vaultize-' + self.ver_no + '-AD-Setup.exe size is ok')
        else:
            print('Vaultize-' + self.ver_no + '-AD-Setup.exe size is not in range..')
            self.fail = True

    def Digital_signature_Check(self,filename):
        command = "powershell Get-AuthenticodeSignature " +filename+ "> check_signature.txt"
        os.system(command)
        f = open("check_signature.txt")
        data = f.read()
        if "Valid" in data:
            print(filename+" digitally signed..")
        else:
            print(filename+" Not digitally signed..")

    def check_digital_signature_Vaultize(self):

        print("checking digital signature for Vaultize Agent....")
        Vaultize_setup_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Setup.exe')
        self.Digital_signature_Check(Vaultize_setup_exe)

        Vaultize_silent = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Silent-Setup.exe')
        self.Digital_signature_Check(Vaultize_silent)

    def check_digital_signature_Outlook_addin(self):
        print("checking digital signature for Vaultize Outlook Addin....")
        Outlook_addin_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Outlook-Addin-Installer.exe')
        self.Digital_signature_Check(Outlook_addin_exe)

        Outlook_addin_msi = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Outlook-Addin-Installer.msi')
        self.Digital_signature_Check(Outlook_addin_msi)

        Outlook_addin64_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Outlook-Addin-x64-Installer.exe')
        self.Digital_signature_Check(Outlook_addin64_exe)

        Outlook_addin32_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Outlook-Addin-x86-Installer.exe')
        self.Digital_signature_Check(Outlook_addin32_exe)

        Outlook_addin64_msi = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Outlook-Addin-x64-Installer.msi')
        self.Digital_signature_Check(Outlook_addin64_msi)

        Outlook_addin32_msi = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Outlook-Addin-x86-Installer.msi')
        self.Digital_signature_Check(Outlook_addin32_msi)

    def check_digital_signature_vDRM(self):
        print("checking digital signature for vDRM Agent....")
        vDRM_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-vDRM-Setup.exe')
        self.Digital_signature_Check(vDRM_exe)

        vDRM_msi = os.path.join(self.path,'Vaultize-' + self.ver_no + '-vDRM-Setup.msi')
        self.Digital_signature_Check(vDRM_msi)

        vDRM_Standalone_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-vDRM-Standalone.exe')
        self.Digital_signature_Check(vDRM_Standalone_exe)

    def check_digital_signature_AD_Notes_plugin(self):
        Notes_plugin_setup = os.path.join(self.path,'Vaultize-' + self.ver_no + '-NotesPlugin-Setup.exe')
        self.Digital_signature_Check(Notes_plugin_setup)

        Vaultize_AD_setup_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-AD-Setup.exe')
        self.Digital_signature_Check(Vaultize_AD_setup_exe)

        Vaultize_AD_setup_msi = os.path.join(self.path,'Vaultize-' + self.ver_no + '-AD-Setup.msi')
        self.Digital_signature_Check(Vaultize_AD_setup_msi)

    def check_file_type_and_bitness_Vaultize(self):
        #check file type and bitness
        Vaultize_setup_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Setup.exe')
        result = magic.from_file(Vaultize_setup_exe)
        print(result)
        if "PE32" in result:
            print("Vaultize-"+self.ver_no+"-setup.exe is Valid Executable(.exe) File.")
            if "PE32+" in result:
                print("Vaultize-"+self.ver_no+"-setup.exe is 64-bit.")
            else:
                print("Vaultize-"+self.ver_no+"-setup.exe is 32-bit.")
        else:
            print("file is not Valid executable(.exe).. ",Vaultize_setup_exe)
            self.fail = True


        Vaultize_silent = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Silent-Setup.exe')
        result = magic.from_file(Vaultize_silent)
        print(result)
        if "PE32" in result:
            print("Vaultize-" + self.ver_no + "-Silent-Setup.exe is Valid Executable(.exe) File.")
            if "PE32+" in result:
                print("Vaultize-" + self.ver_no + "-Silent-Setup.exe is 64-bit.")
            else:
                print("Vaultize-" + self.ver_no + "-Silent-Setup.exe is 32-bit.")
        else:
            print("file is not valid executable(.exe) ",Vaultize_silent)
            self.fail = True


    def check_file_type_and_bitness_Outlook_Addin(self):
        Outlook_addin_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Outlook-Addin-Installer.exe')
        result = magic.from_file(Outlook_addin_exe)
        print(result)
        if "PE32" in result:
            print("Vaultize-" + self.ver_no + "-Outlook-Addin-Installer.exe is Valid Executable(.exe) File.")
            if "PE32+" in result:
                print("Vaultize-" + self.ver_no + "-Outlook-Addin-Installer.exe is 64-bit.")
            else:
                print("Vaultize-" + self.ver_no + "-Outlook-Addin-Installer.exe is 32-bit.")
        else:
            print("file is not valid executable(.exe) ",Outlook_addin_exe)
            self.fail = True

        Outlook_addin32_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Outlook-Addin-x86-Installer.exe')
        result = magic.from_file(Outlook_addin32_exe)
        print(result)
        if "PE32" in result:
            print("Vaultize-" + self.ver_no + "-Outlook-Addin-x86-Installer.exe is Valid Executable(.exe) File.")
            if "PE32+" in result:
                print("Vaultize-" + self.ver_no + "-Outlook-Addin-x86-Installer.exe is 64-bit.")
            else:
                print("Vaultize-" + self.ver_no + "-Outlook-Addin-x86-Installer.exe is 32-bit.")
        else:
            print("file is not valid executable(.exe) ",Outlook_addin32_exe)
            self.fail = True

        Outlook_addin64_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Outlook-Addin-x64-Installer.exe')
        result = magic.from_file(Outlook_addin64_exe)
        print(result)
        if "PE32" in result:
            print("Vaultize-" + self.ver_no + "-Outlook-Addin-x64-Installer.exe is Valid Executable(.exe) File.")
            if "PE32+" in result:
                print("Vaultize-" + self.ver_no + "-Outlook-Addin-x64-Installer.exe is 64-bit.")
            else:
                print("Vaultize-" + self.ver_no + "-Outlook-Addin-x64-Installer.exe is 32-bit.")
        else:
            print("file is not Valid executable(.exe).. ",Outlook_addin64_exe)
            self.fail = True

    def check_file_type_and_bitness_vDRM(self):
        vDRM_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-vDRM-Setup.exe')
        result = magic.from_file(vDRM_exe)
        print(result)
        if "PE32" in result:
            print("Vaultize-" + self.ver_no + "-vDRM-Setup.exe is Valid Executable(.exe) File.")
            if "PE32+" in result:
                print("Vaultize-" + self.ver_no + "-vDRM-Setup.exe is 64-bit.")
            else:
                print("Vaultize-" + self.ver_no + "-vDRM-Setup.exe is 32-bit.")
        else:
            print("file is not Valid executable(.exe).. ",vDRM_exe)
            self.fail = True

        vDRM_Standalone_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-vDRM-Standalone.exe')
        result = magic.from_file(vDRM_Standalone_exe)
        print(result)
        if "PE32" in result:
            print("Vaultize-" + self.ver_no + "-vDRM-Standalone.exe is Valid Executable(.exe) File.")
            if "PE32+" in result:
                print("Vaultize-" + self.ver_no + "-vDRM-Standalone.exe is 64-bit.")
            else:
                print("Vaultize-" + self.ver_no + "-vDRM-Standalone.exe is 32-bit.")
        else:
            print("file is not Valid executable(.exe).. ",vDRM_Standalone_exe)
            self.fail = True

    def check_file_type_and_bitness_AD_Notes(self):
        Notes_plugin_setup = os.path.join(self.path,'Vaultize-' + self.ver_no + '-NotesPlugin-Setup.exe')
        result = magic.from_file(Notes_plugin_setup)
        print(result)
        if "PE32" in result:
            print("Vaultize-" + self.ver_no + "-NotesPlugin-Setup.exe is Valid Executable(.exe) File.")
            if "PE32+" in result:
                print("Vaultize-" + self.ver_no + "-NotesPlugin-Setup.exe is 64-bit.")
            else:
                print("Vaultize-" + self.ver_no + "-NotesPlugin-Setup.exe is 32-bit.")
        else:
            print("file is not Valid executable(.exe).. ",Notes_plugin_setup)
            self.fail = True

        Vaultize_AD_setup_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-AD-Setup.exe')
        result = magic.from_file(Vaultize_AD_setup_exe)
        print(result)
        if "PE32" in result:
            print("Vaultize-" + self.ver_no + "-AD-Setup.exe is Valid Executable(.exe) File.")
            if "PE32+" in result:
                print("Vaultize-" + self.ver_no + "-AD-Setup.exe is 64-bit.")
            else:
                print("Vaultize-" + self.ver_no + "-AD-Setup.exe is 32-bit.")
        else:
            print("file is not Valid executable(.exe).. ",Vaultize_AD_setup_exe)
            self.fail = True

    def check_file_name_pattern(self):

        dir_list = os.listdir(self.path)
        current_year = datetime.now().strftime('%y')  # current year without century
        for file in dir_list:
            self.dflag=0
            self.mflag=0
            self.yflag=0
            pattern = "[0-9][0-9].[0-9][0-9].[0-9][0-9]"
            match = re.search(pattern, file)
            data = match.group()
            list = data.split(".")
            yy = list[0]
            mm = list[1]
            dd = list[2]

            if yy != current_year:
                print("filename format is not correct at YY:: "+file)
                self.fail=True
            else:
                self.yflag=1

            if dd > 31 and dd < 1:
                print("filename format is not correct at DD")
                self.fail=True
            else:
                self.dflag=1

            if mm > 12 and mm < 1:
                print("filename format is not correct at MM")
                self.fail=True
            else:
                self.mflag=1

            date = yy + "." + mm + "." + dd
            if self.yflag==1 and self.mflag==1 and self.dflag==1 and date in file:
                print("file name pattern is correct for:: "+file)
                print("file name pattern is:: "+date)

    def check_Timestamp_Vaultize(self):
        Vaultize_setup_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Setup.exe')
        print("timestamp for Vaultize-"+ self.ver_no + '-Setup.exe')
        modified_date_time = os.path.getmtime(Vaultize_setup_exe)
        print "Modified timestamp:: " +datetime.fromtimestamp(modified_date_time).strftime('%y-%m-%d %H:%M:%S')
        created_date_time = os.path.getctime(Vaultize_setup_exe)
        print "Createded timestamp:: "+datetime.fromtimestamp(created_date_time).strftime('%y-%m-%d %H:%M:%S')+"\n"

        Vaultize_silent = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Silent-Setup.exe')
        print("timestamp for Vaultize-" + self.ver_no + '-Silent-Setup.exe')
        modified_date_time = os.path.getmtime(Vaultize_silent)
        print "Modified timestamp:: " + datetime.fromtimestamp(modified_date_time).strftime('%y-%m-%d %H:%M:%S')
        created_date_time = os.path.getctime(Vaultize_silent)
        print "Createded timestamp:: " + datetime.fromtimestamp(created_date_time).strftime('%y-%m-%d %H:%M:%S') + "\n"

    def check_Timestamp_Outlook_Addin(self):
        Outlook_addin_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Outlook-Addin-Installer.exe')
        print("timestamp for Vaultize-" + self.ver_no + '-Outlook-Addin-Installer.exe')
        modified_date_time = os.path.getmtime(Outlook_addin_exe)
        print "Modified timestamp:: " + datetime.fromtimestamp(modified_date_time).strftime('%y-%m-%d %H:%M:%S')
        created_date_time = os.path.getctime(Outlook_addin_exe)
        print "Createded timestamp:: " + datetime.fromtimestamp(created_date_time).strftime('%y-%m-%d %H:%M:%S') + "\n"

        Outlook_addin_msi = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Outlook-Addin-Installer.msi')
        print("timestamp for Vaultize-" + self.ver_no + '-Outlook-Addin-Installer.msi')
        modified_date_time = os.path.getmtime(Outlook_addin_msi)
        print "Modified timestamp:: " + datetime.fromtimestamp(modified_date_time).strftime('%y-%m-%d %H:%M:%S')
        created_date_time = os.path.getctime(Outlook_addin_msi)
        print "Createded timestamp:: " + datetime.fromtimestamp(created_date_time).strftime('%y-%m-%d %H:%M:%S') + "\n"

        Outlook_addin64_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Outlook-Addin-x64-Installer.exe')
        print("timestamp for Vaultize-" + self.ver_no + '-Outlook-Addin-x64-Installer.exe')
        modified_date_time = os.path.getmtime(Outlook_addin64_exe)
        print "Modified timestamp:: " + datetime.fromtimestamp(modified_date_time).strftime('%y-%m-%d %H:%M:%S')
        created_date_time = os.path.getctime(Outlook_addin64_exe)
        print "Createded timestamp:: " + datetime.fromtimestamp(created_date_time).strftime('%y-%m-%d %H:%M:%S') + "\n"

        Outlook_addin32_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Outlook-Addin-x86-Installer.exe')
        print("timestamp for Vaultize-" + self.ver_no + '-Outlook-Addin-x86-Installer.exe')
        modified_date_time = os.path.getmtime(Outlook_addin32_exe)
        print "Modified timestamp:: " + datetime.fromtimestamp(modified_date_time).strftime('%y-%m-%d %H:%M:%S')
        created_date_time = os.path.getctime(Outlook_addin32_exe)
        print "Createded timestamp:: " + datetime.fromtimestamp(created_date_time).strftime('%y-%m-%d %H:%M:%S') + "\n"

        Outlook_addin64_msi = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Outlook-Addin-x64-Installer.msi')
        print("timestamp for Vaultize-" + self.ver_no + '-Outlook-Addin-x64-Installer.msi')
        modified_date_time = os.path.getmtime(Outlook_addin64_msi)
        print "Modified timestamp:: " + datetime.fromtimestamp(modified_date_time).strftime('%y-%m-%d %H:%M:%S')
        created_date_time = os.path.getctime(Outlook_addin64_msi)
        print "Createded timestamp:: " + datetime.fromtimestamp(created_date_time).strftime('%y-%m-%d %H:%M:%S') + "\n"

        Outlook_addin32_msi = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Outlook-Addin-x86-Installer.msi')
        print("timestamp for Vaultize-" + self.ver_no + '-Outlook-Addin-x86-Installer.msi')
        modified_date_time = os.path.getmtime(Outlook_addin32_msi)
        print "Modified timestamp:: " + datetime.fromtimestamp(modified_date_time).strftime('%y-%m-%d %H:%M:%S')
        created_date_time = os.path.getctime(Outlook_addin32_msi)
        print "Createded timestamp:: " + datetime.fromtimestamp(created_date_time).strftime('%y-%m-%d %H:%M:%S') + "\n"

    def check_Timestamp_vDRM(self):
        vDRM_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-vDRM-Setup.exe')
        print("timestamp for Vaultize-" + self.ver_no + '-vDRM-Setup.exe')
        modified_date_time = os.path.getmtime(vDRM_exe)
        print "Modified timestamp:: " + datetime.fromtimestamp(modified_date_time).strftime('%y-%m-%d %H:%M:%S')
        created_date_time = os.path.getctime(vDRM_exe)
        print "Createded timestamp:: " + datetime.fromtimestamp(created_date_time).strftime('%y-%m-%d %H:%M:%S') + "\n"

        vDRM_msi = os.path.join(self.path,'Vaultize-' + self.ver_no + '-vDRM-Setup.msi')
        print("timestamp for Vaultize-" + self.ver_no + '-vDRM-Setup.msi')
        modified_date_time = os.path.getmtime(vDRM_msi)
        print "Modified timestamp:: " + datetime.fromtimestamp(modified_date_time).strftime('%y-%m-%d %H:%M:%S')
        created_date_time = os.path.getctime(vDRM_msi)
        print "Createded timestamp:: " + datetime.fromtimestamp(created_date_time).strftime('%y-%m-%d %H:%M:%S') + "\n"

        vDRM_Standalone_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-vDRM-Standalone.exe')
        print("timestamp for Vaultize-" + self.ver_no + '-vDRM-Standalone.exe')
        modified_date_time = os.path.getmtime(vDRM_Standalone_exe)
        print "Modified timestamp:: " + datetime.fromtimestamp(modified_date_time).strftime('%y-%m-%d %H:%M:%S')
        created_date_time = os.path.getctime(vDRM_Standalone_exe)
        print "Createded timestamp:: " + datetime.fromtimestamp(created_date_time).strftime('%y-%m-%d %H:%M:%S') + "\n"

    def check_Timestamp_AD_Notes(self):
        Notes_plugin_setup=os.path.join(self.path,'Vaultize-'+self.ver_no+'-NotesPlugin-Setup.exe')
        print("timestamp for Vaultize-" + self.ver_no + '-NotesPlugin-Setup.exe')
        modified_date_time = os.path.getmtime(Notes_plugin_setup)
        print "Modified timestamp:: " + datetime.fromtimestamp(modified_date_time).strftime('%y-%m-%d %H:%M:%S')
        created_date_time = os.path.getctime(Notes_plugin_setup)
        print "Createded timestamp:: " + datetime.fromtimestamp(created_date_time).strftime('%y-%m-%d %H:%M:%S') + "\n"

        Vaultize_AD_setup_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-AD-Setup.exe')
        print("timestamp for Vaultize-" + self.ver_no + '-AD-Setup.exe')
        modified_date_time = os.path.getmtime(Vaultize_AD_setup_exe)
        print "Modified timestamp:: " + datetime.fromtimestamp(modified_date_time).strftime('%y-%m-%d %H:%M:%S')
        created_date_time = os.path.getctime(Vaultize_AD_setup_exe)
        print "Createded timestamp:: " + datetime.fromtimestamp(created_date_time).strftime('%y-%m-%d %H:%M:%S') + "\n"

        Vaultize_AD_setup_msi=os.path.join(self.path,'Vaultize-' + self.ver_no + '-AD-Setup.msi')
        print("timestamp for Vaultize-" + self.ver_no + '-AD-Setup.msi')
        modified_date_time = os.path.getmtime(Vaultize_AD_setup_msi)
        print "Modified timestamp:: " + datetime.fromtimestamp(modified_date_time).strftime('%y-%m-%d %H:%M:%S')
        created_date_time = os.path.getctime(Vaultize_AD_setup_msi)
        print "Createded timestamp:: " + datetime.fromtimestamp(created_date_time).strftime('%y-%m-%d %H:%M:%S') + "\n"

    def validate_build_number_Vaultize(self):
        print("validating build number..")
        Vaultize_setup_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Setup.exe')
        res=self.get_file_info(Vaultize_setup_exe,'FileVersion')
        print(res)
        if self.ver_no in res:
            print("Version number of Vaultize-" + self.ver_no + '-Setup.exe is correct' )
        else:
            print("Version number of Vaultize-" + self.ver_no + '-Setup.exe is Not assigned or Not Matched...')

        Vaultize_silent = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Silent-Setup.exe')
        res = self.get_file_info(Vaultize_silent, 'FileVersion')
        print(res)
        if self.ver_no in res:
            print("Version number of Vaultize-" + self.ver_no + '-Silent-Setup.exe is correct')
        else:
            print("Version number of Vaultize-" + self.ver_no + '-Silent-Setup.exe is Not assigned or Not Matched...')

    def validate_build_number_Outlook_Addin(self):
        Outlook_addin_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Outlook-Addin-Installer.exe')
        res = self.get_file_info(Outlook_addin_exe, 'FileVersion')
        print(res)
        if self.ver_no in res:
            print("Version number of Vaultize-" + self.ver_no + '-Outlook-Addin-Installer.exe is correct')
        else:
            print("Version number of Vaultize-" + self.ver_no + '-Outlook-Addin-Installer.exe is Not assigned or Not Matched...')

        Outlook_addin_msi = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Outlook-Addin-Installer.msi')
        res = self.get_file_info(Outlook_addin_msi, 'FileVersion')
        print(res)
        if self.ver_no in res:
            print("Version number of Vaultize-" + self.ver_no + '-Outlook-Addin-Installer.msi is correct')
        else:
            print("Version number of Vaultize-" + self.ver_no + '-Outlook-Addin-Installer.msi is Not assigned or Not Matched...')

        Outlook_addin64_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Outlook-Addin-x64-Installer.exe')
        res = self.get_file_info(Outlook_addin64_exe, 'FileVersion')
        print(res)
        if self.ver_no in res:
            print("Version number of Vaultize-" + self.ver_no + '-Outlook-Addin-x64-Installer.exe is correct')
        else:
            print("Version number of Vaultize-" + self.ver_no + '-Outlook-Addin-x64-Installer.exe is Not assigned or Not Matched...')

        Outlook_addin32_msi = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Outlook-Addin-x86-Installer.msi')
        res = self.get_file_info(Outlook_addin32_msi, 'FileVersion')
        print(res)
        if self.ver_no in res:
            print("Version number of Vaultize-" + self.ver_no + '-Outlook-Addin-x86-Installer.msi is correct')
        else:
            print("Version number of Vaultize-" + self.ver_no + '-Outlook-Addin-x86-Installer.msi is Not assigned or Not Matched...')

        Outlook_addin64_msi = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Outlook-Addin-x64-Installer.msi')
        res = self.get_file_info(Outlook_addin64_msi, 'FileVersion')
        print(res)
        if self.ver_no in res:
            print("Version number of Vaultize-" + self.ver_no + '-Outlook-Addin-x64-Installer.msi is correct')
        else:
            print("Version number of Vaultize-" + self.ver_no + '-Outlook-Addin-x64-Installer.msi is Not assigned or Not Matched...')

        Outlook_addin32_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-Outlook-Addin-x86-Installer.exe')
        res = self.get_file_info(Outlook_addin32_exe, 'FileVersion')
        print(res)
        if self.ver_no in res:
            print("Version number of Vaultize-" + self.ver_no + '-Outlook-Addin-x86-Installer.exe is correct')
        else:
            print("Version number of Vaultize-" + self.ver_no + '-Outlook-Addin-x86-Installer.exe is Not assigned or Not Matched...')

    def validate_build_number_vDRM(self):
        vDRM_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-vDRM-Setup.exe')
        res = self.get_file_info(vDRM_exe, 'FileVersion')
        print(res)
        if self.ver_no in res:
            print("Version number of Vaultize-" + self.ver_no + '-vDRM-Setup.exe is correct')
        else:
            print("Version number of Vaultize-" + self.ver_no + '-vDRM-Setup.exe is Not assigned or Not Matched...')

        vDRM_msi = os.path.join(self.path,'Vaultize-' + self.ver_no + '-vDRM-Setup.msi')
        res = self.get_file_info(vDRM_msi, 'FileVersion')
        print(res)
        if self.ver_no in res:
            print("Version number of Vaultize-" + self.ver_no + '-vDRM-Setup.msi is correct')
        else:
            print("Version number of Vaultize-" + self.ver_no + '-vDRM-Setup.msi is Not assigned or Not Matched...')

        vDRM_Standalone_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-vDRM-Standalone.exe')
        res = self.get_file_info(vDRM_Standalone_exe, 'FileVersion')
        print(res)
        if self.ver_no in res:
            print("Version number of Vaultize-" + self.ver_no + '-vDRM-Standalone.exe is correct')
        else:
            print("Version number of Vaultize-" + self.ver_no + '-vDRM-Standalone.exe is Not assigned or Not Matched...')

    def validate_build_number_AD_Notes(self):
        Notes_plugin_setup = os.path.join(self.path,'Vaultize-' + self.ver_no + '-NotesPlugin-Setup.exe')
        res = self.get_file_info(Notes_plugin_setup, 'FileVersion')
        print(res)
        if self.ver_no in res:
            print("Version number of Vaultize-" + self.ver_no + '-NotesPlugin-Setup.exe is correct')
        else:
            print("Version number of Vaultize-" + self.ver_no + '-NotesPlugin-Setup.exe is Not assigned or Not Matched...')

        Vaultize_AD_setup_exe = os.path.join(self.path,'Vaultize-' + self.ver_no + '-AD-Setup.exe')
        res = self.get_file_info(Vaultize_AD_setup_exe, 'FileVersion')
        print(res)
        if self.ver_no in res:
            print("Version number of Vaultize-" + self.ver_no + '-AD-Setup.exe is correct')
        else:
            print("Version number of Vaultize-" + self.ver_no + '-AD-Setup.exe is Not assigned or Not Matched...')

        Vaultize_AD_setup_msi = os.path.join(self.path,'Vaultize-' + self.ver_no + '-AD-Setup.msi')
        res = self.get_file_info(Vaultize_AD_setup_msi, 'FileVersion')
        print(res)
        if self.ver_no in res:
            print("Version number of Vaultize-" + self.ver_no + '-AD-Setup.msi is correct')
        else:
            print("Version number of Vaultize-" + self.ver_no + '-AD-Setup.msi is Not assigned or Not Matched...')

    def get_file_info(self,filename, info):
        size = windll.version.GetFileVersionInfoSizeA(filename, None)
        if not size:
            return ''
        res = create_string_buffer(size)
        windll.version.GetFileVersionInfoA(filename, None, size, res)
        r = c_uint()
        l = c_uint()
        windll.version.VerQueryValueA(res, '\\VarFileInfo\\Translation',
                                      byref(r), byref(l))
        if not l.value:
            return ''
        codepages = array.array('H', string_at(r.value, l.value))
        codepage = tuple(codepages[:2].tolist())
        windll.version.VerQueryValueA(res, ('\\StringFileInfo\\%04x%04x\\'
                                            + info) % codepage, byref(r), byref(l))
        return string_at(r.value,l.value).strip()

    def cleanup(self):
        print "Inside cleanup"
        #delete

def main(*args, **kwargs):
    if len(sys.argv) != 4:
        print("USAGE:")
        print("python __init__.py --path=<Dist_path> --ver=<version number> --check=<vaultize_agent/outlook_addin/vDRM/ALL>")
        os._exit(1)

    sys.argv.pop(0)
    path = None
    agent = None
    ver_no = None

    for arg in sys.argv:
        print(arg)

        if arg.find(OPTION_CHECK_AGENT) != -1:
            agent = arg.split("=")[1]
            #print serverzip
        if arg.find(OPTION_VER_NO) != -1 :
            ver_no = arg.split("=")[1]
            #print ver_no

        if arg.find(OPTION_DIST_PATH) != -1 :
            path = arg.split("=")[1]

    print("Validating for:: "+agent)
    if agent not in ("Vaultize_Agent", "Outlook_Addin", "vDRM", "ALL"):
        print("choose correct option from <Vaultize_Agent/Outlook_Addin/vDRM/ALL>...")
        os.exit(1)
    if not ver_no  or not agent or not path:
        print("USAGE:")
        print("python __init__.py --path=<Dist_path> --ver=<version number> --check<Vaultize_Agent/Outlook_Addin/vDRM/ALL>")
        os._exit(1)
    exit_code=Agent_Build_Validator(path, ver_no, agent).run()
    print "Quitting with exit code (True==Failed, False==Pass): ", exit_code
if __name__ == "__main__":
    try:
        main().run()
    except:
        pass