from pathlib import Path, WindowsPath
import subprocess
from shutil import copy2
import sys
import os
from win32com.client import Dispatch
import win32serviceutil
import time

def display_uninstall_status(func):
    def call_func(self,*args, **kwargs):
        uninstall_name = kwargs.get('uninstall_name',"")
        sys.stdout.write("Uninstalling %s" % uninstall_name)
        sys.stdout.flush()
        retval=func(self,*args,**kwargs)
        sys.stdout.write("\rUninstall %s: %s\n" % (uninstall_name,retval))
        sys.stdout.flush()
        return retval
    return call_func

def display_install_status(func):
    def call_func(self,*args, **kwargs):
        install_name = kwargs.get('install_name',"")
        sys.stdout.write("Installing %s" % install_name)
        sys.stdout.flush()
        retval=func(self,*args,**kwargs)
        sys.stdout.write("\rInstall %s: %s\n" % (install_name,retval))
        sys.stdout.flush()
        return retval
    return call_func


class RemoteInstall(object):
    def __init__(self,comp_name,script_path):
        self.comp_name = comp_name
        self.script_path = script_path

    def execute_remote(self,exe,exe_params=[],remote_params=[],verbose=False):
        """
        Execute a remote executable and return verbose output if specified
        Args:
            exe: full executable path on remote machine
        Kwargs:
            exe_params: arguments for supplied exe
            remote_params: arguments specifically for psexec
            verbose: Specify whether to print console text of command
        """
        psexec_path = self.script_path.parent.joinpath('psexec.exe')
        if psexec_path.exists():

            comp_name_psexec = r"\\%s" % self.comp_name

            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE

            popen_list = [str(x) for x in [str(psexec_path),"-accepteula"] + remote_params + [comp_name_psexec,str(exe)] + exe_params]
            result = subprocess.Popen(popen_list,stdout=subprocess.PIPE,stderr=subprocess.PIPE,stdin=subprocess.PIPE,shell=False,startupinfo=startupinfo)
            out1, out1_err = result.communicate()
            if verbose:
                print(out1.decode("utf-8"))
                print(out1_err.decode("utf-8"))
            return result.returncode
        else:
            print("\r*****************Missing PSEXEC*****************")
            return 2

    def execute_other_remote_psexec(self,exe_params=[],remote_params=[],verbose=False):
        """
        Execute a remote command through psexec and return verbose output if specified
        Kwargs:
            exe_params: command to run and arguments
            remote_params: arguments specifically for psexec
            verbose: Specify whether to print console text of command
        """
        psexec_path = self.script_path.parent.joinpath('psexec.exe')
        if psexec_path.exists():

            comp_name_psexec = r"\\%s" % self.comp_name

            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE

            popen_list = [str(x) for x in [str(psexec_path),"-accepteula"] + remote_params + [comp_name_psexec] + exe_params]
            result = subprocess.Popen(popen_list,stdout=subprocess.PIPE,stderr=subprocess.PIPE,stdin=subprocess.PIPE,shell=False,startupinfo=startupinfo)
            out1, out1_err = result.communicate()
            if verbose:
                print(out1.decode("utf-8"))
                print(out1_err.decode("utf-8"))
            return result.returncode
        else:
            print("\r*****************Missing PSEXEC*****************")
            return 2

    def execute_other_remote(self,exe_params=[],verbose=False):
        """
        Execute command given in exe_params (no psexec)
        Kwargs:
            exe_params: command to run and arguments
            verbose: Specify whether to print console text of command
        """
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE

        result = subprocess.Popen(exe_params, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                    stdin=subprocess.PIPE, shell=False, startupinfo=startupinfo)
        out1, out1_err = result.communicate()
        if verbose:
            print(out1.decode("utf-8"))
            print(out1_err.decode("utf-8"))
        return result.returncode

    @display_uninstall_status
    def uninstall_msi(self, product_code, uninstall_name="", verbose=False):
        """
        Standard silent no restart uninstall MSI option using product code
        args:
            product_code: msi installer code of application
        Kwargs:
            uninstall_name: visible name used for console output
            verbose: Specify whether to print console text of command
        """
        return self.execute_remote("msiexec.exe", exe_params=["/x", product_code, "/qn", "/norestart"], remote_params=["-s", "-i"], verbose=verbose)

    @display_uninstall_status
    def uninstall_exe(self, uninstall_list, uninstall_name="", verbose=False):
        """
        Uninstall using an executable and arguments
        Args:
            uninstall_list: exe and uninstall arguments
        Kwargs:
            uninstall_name: visible name used for console output
            verbose: Specify whether to print console text of command
        """
        return self.execute_remote(uninstall_list[0], exe_params=uninstall_list[1:], remote_params=["-s"], verbose=verbose)

    @display_install_status
    def install_msi(self, destination_dir, install_name="", optional_params= [], remote_params=[], copy_file=True, verbose=False):
        """
        Copy msi file to remote pc and execute install
        args:
            destination_dir: where to put the msi file on remote PC
        kwargs:
            install_name: visible name used for console output
            optional_params: arguments for msi
            remote_params: arguments specifically for psexec
            copy_file: Specify whether to copy msi to remote PC (used to execute msi that may already exist on PC)
            verbose: Specify whether to print console text of command
        """
        if copy_file:
            source_file = self.script_path.parent.joinpath(install_name)
            destination_dir = Path(destination_dir)
            destination_dir.mkdir(parents=True,exist_ok=True)

            copy2(source_file,destination_dir)

        destination_file = destination_dir.joinpath(install_name)
        if destination_file.exists():
            return self.execute_remote("msiexec.exe", exe_params=["/i", destination_file, "/qn", "/norestart"] + optional_params, remote_params=["-s"] + remote_params, verbose=verbose)
        else:
            return 1

    @display_install_status
    def apply_reg(self, destination_dir, install_name="", verbose=False):
        """
        Copy reg file to remote pc and execute
        args:
            destination_dir: where to put the reg file on remote PC
        kwargs:
            install_name: visible name used for console output
            verbose: Specify whether to print console text of command
        """
        source_file = self.script_path.parent.joinpath(install_name)
        destination_dir = Path(destination_dir)
        destination_dir.mkdir(parents=True,exist_ok=True)

        copy2(source_file,destination_dir)
        destination_file = destination_dir.joinpath(install_name)
        if destination_file.exists():
            return self.execute_remote("regedit.exe", exe_params=["/s", destination_file], remote_params=["-s"], verbose=verbose)
        else:
            return 1

    @display_install_status
    def install_msi_copy_and_params(self, destination_dir, install_name="", extra_files=[], extra_params=[], verbose=False):
        """
        Copy msi file and other required files to remote pc and execute install
        args:
            destination_dir: where to put the msi file on remote PC
        kwargs:
            install_name: visible name used for console output
            extra_files: files to copy with msi
            extra_params: arguments for msi
            verbose: Specify whether to print console text of command
        """
        abort = False

        source_file = self.script_path.parent.joinpath(install_name)
        destination_dir = Path(destination_dir)
        destination_dir.mkdir(parents=True,exist_ok=True)
        copy2(source_file,destination_dir)

        for f in extra_files:
            source_f = self.script_path.parent.joinpath(f)
            copy2(source_f,destination_dir)
            if not destination_dir.joinpath(f).exists():
                abort = True

        destination_file = destination_dir.joinpath(install_name)
        msi_params = ["/i",str(destination_file)] + extra_params
        #msi_params.append("/L*V")
        #msi_params.append("C:\\dell\\dragonmedicalone\\nuancedmo.log")

        if destination_file.exists() and abort == False:
            return self.execute_remote("msiexec.exe", exe_params=msi_params, remote_params=["-s"], verbose=verbose)
        else:
            return 1

    @display_install_status
    def install_exe(self,destination_dir,parameters=[],install_name="",copy_file=True,verbose=False):
        """
        Copy msi file and other required files to remote pc and execute install
        args:
            destination_dir: where to put the msi file on remote PC
        kwargs:
            parameters: arguments for exe
            install_name: visible name used for console output
            copy_file: Specify whether to copy exe to remote PC (used to execute exe that may already exist on PC)
            verbose: Specify whether to print console text of command
        """
        if copy_file:
            source_file = self.script_path.parent.joinpath(install_name)
            destination_dir = Path(destination_dir)
            destination_dir.mkdir(parents=True,exist_ok=True)

            copy2(source_file,destination_dir)
        destination_file = destination_dir.joinpath(install_name)
        if destination_file.exists():
            return self.execute_remote(destination_file,exe_params=parameters,remote_params=[],verbose=verbose)
        else:
            return 1

    @display_uninstall_status
    def uninstall_exe_copy(self, destination_dir, parameters=[], uninstall_name="", verbose=False):
        """
        Uninstall using an executable and arguments
        Args:
            destination_dir: directory to put uninstall exe
        Kwargs:
            parameters: uninstall arguments for exe
            uninstall_name: visible name used for console output
            verbose: Specify whether to print console text of command
        """
        source_file = self.script_path.parent.joinpath(uninstall_name)
        destination_dir = Path(destination_dir)
        destination_dir.mkdir(parents=True,exist_ok=True)

        copy2(source_file,destination_dir)
        destination_file = destination_dir.joinpath(uninstall_name)
        if destination_file.exists():
            return self.execute_remote(destination_file, exe_params=parameters, remote_params=[], verbose=verbose)
        else:
            return 1

    def add_icons(self,profile_path,icon_name,target_path):
        """
        Create Icons to be placed in profile_path
        args:
            profile_path: path to location to store icon
            icon_name: name of icon file
            target_path: target icon points to
        """
        ret = 5
        if Path(profile_path).exists():
            try:
                shortcut_path = Path(profile_path).joinpath(icon_name + ".lnk")
                shell = Dispatch('WScript.Shell')
                shortcut = shell.CreateShortcut(str(shortcut_path))
                shortcut.Targetpath = str(Path(target_path))
                shortcut.IconLocation = str(Path(target_path))
                shortcut.save()
                ret = 0
            except Exception as e:
                print(e)
                ret = -1
            finally:
                return ret
        else:
            return ret

    def copy_icons(self,profile_path,icon_name):
        """
        Copy icon file to location
        args:
            profile_path: path to place icon
            icon_name: filename of icon
        """
        if Path(profile_path).exists():
            try:
                source_file = self.script_path.parent.joinpath(icon_name)
                copy2(source_file,profile_path)
                destination_file = profile_path.joinpath(icon_name)
                if destination_file.exists():
                    return 0
                else:
                    return 1
            except Exception as e:
                print(e)
                return -1
        else:
            return 5

    def copy_file(self,source_name,destination_dir):
        """
        Copy a file to a destination and check
        args:
            source_name: source filename (in same directory as calling script)
            destination_dir: destination to place file
        """
        source_file = self.script_path.parent.joinpath(source_name)
        if Path(source_file).exists():
            try:
                copy2(source_file, destination_dir)
                destination_file = destination_dir.joinpath(source_name)
                if destination_file.exists():
                    return 0
                else:
                    return 1
            except Exception as e:
                print(e)
                return -1
        else:
            return 5

    def remove_url_icons(self,profile_path,url):
        """
        Reads .url file in directory and removes file with matching URL
        args:
            profile_path: directory to check for URL
            url: URL to check for
        """
        files_to_delete = []
        found = False
        if Path(profile_path).exists() and Path(profile_path).is_dir():
            for x in Path(profile_path).iterdir():
                if x.is_file() and x.suffix == ".url":
                    with open(str(x),"r") as urlfile:
                        if url in urlfile.read():
                            files_to_delete.append(x)
            for x in files_to_delete:
                found = True
                x.unlink()
            if found:
                return 0
            else:
                return 5
        else:
            return 5

    def unzip(self,destination_dir,sourcezip):
        """
        Using unzip application stored in calling script path unzip zip to directory
        http://infozip.sourceforge.net/
        args:
            destination_dir: destination to place contents
            sourcezip: zip to extract
        """
        s_unzip_exe = self.script_path.parent.joinpath("unzip.exe")
        s_unzip_dll = self.script_path.parent.joinpath("unzip32.dll")
        destination_dir = Path(destination_dir)
        destination_dir.mkdir(parents=True, exist_ok=True)

        copy2(s_unzip_exe,destination_dir)
        d_unzip_exe = destination_dir.joinpath("unzip.exe")
        if not d_unzip_exe.exists():
            return 1
        else:
            print("Copied unzip.exe")

        copy2(s_unzip_dll,destination_dir)
        d_unzip_dll = destination_dir.joinpath("unzip32.dll")
        if not d_unzip_dll.exists():
            return 1
        else:
            print("Copied unzip32.dll")

        copy2(sourcezip,destination_dir)
        d_zip_file = destination_dir.joinpath(sourcezip.name)
        if not d_zip_file.exists():
            return 1
        else:
            print("Copied Zip file")

        local_unzip_exe = str(d_unzip_exe).replace("\\\\" + self.comp_name + "\\c$","C:")

        #create zip directory
        #zip_drop_folder = destination_dir.joinpath(d_zip_file.stem)
        #zip_drop_folder.mkdir(parents=True,exist_ok=True)

        if destination_dir.exists():
            exe_params = ["-o","-q", str(d_zip_file).replace("\\\\" +
                                                        self.comp_name + "\\c$", "C:"), "-d", str(destination_dir).replace("\\\\" + self.comp_name + "\\c$", "C:")]
            ret = self.execute_remote(local_unzip_exe,exe_params=exe_params,remote_params=[],verbose=True)
            return ret
        else:
            return 1

    def check_service(self,service_name):
        """
        Check if a service is running on remote PC
        """
        _, s_state, _, _, _, _, _ = win32serviceutil.QueryServiceStatus(
            service_name, self.comp_name)
        return s_state

    def start_service(self,service_name,retry_times=3,restart=False):
        """
        Start service on remote PC. Restart if required.
        args:
            service_name:name of service to start
        kwargs:
            retry_times:how many times to attempt to start service
            restart: specify whether to restart service if already running
        """
        s_state = self.check_service(service_name)
        if s_state == 4:
            if restart:
                print("restarting service %s" % service_name)
                self.stop_service(service_name)
                if self.check_service(service_name) == 1:
                    time.sleep(1)
                    return self.start_service(service_name,retry_times=retry_times,restart=False)
            else:
                return 0

        elif s_state == 1:
            print("starting service %s" % service_name)
            win32serviceutil.StartService(service_name,[],self.comp_name)
            if s_state == 4:
                return 0
            elif s_state == 3 or s_state == 2:
                time.sleep(5)
                return self.start_service(service_name,retry_times=retry_times,restart=False)
            elif retry_times > 0:
                print("retrying...")
                return self.start_service(service_name,retry_times=retry_times-1,restart=False)
            else:
                return -5
        else:
            return self.start_service(service_name, retry_times=retry_times-1,restart=False)

    def stop_service(self,service_name,retry_times=3):
        """
        Stop service on remote PC.
        args:
            service_name: name of service to stop
        kwargs:
            retry_times: how many times to attempt to stop service
        """
        s_state = self.check_service(service_name)
        if s_state == 4:
            print("stopping service")
            win32serviceutil.StopService(service_name,self.comp_name)
            s_state = self.check_service(service_name)
            if s_state == 1:
                return 0
            elif s_state == 3 or s_state == 2:
                time.sleep(5)
                return self.stop_service(service_name,retry_times=retry_times)
            elif retry_times > 0:
                print("retrying...")
                return self.stop_service(service_name, retry_times=retry_times-1)
            else:
                return -5
            
        elif s_state == 1:
            print("stopped")
            return 0
        else:
            return self.stop_service(service_name, retry_times=retry_times-1)
