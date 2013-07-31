'''
Copyright (c) 2013 by JustAMan at GitHub

Permission is hereby granted, free of charge, to any person obtaining a copy of
this software and associated documentation files (the "Software"), to deal in
the Software without restriction, including without limitation the rights to
use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
the Software, and to permit persons to whom the Software is furnished to do so,
subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

This file provides the ability to re-elevate the rights of running Python script to
Administrative ones allowing console of elevated process to remain user-interactive.

You should use elevateAdminRights() at the very possible start point of your scripts because
it will re-launch your script from the very beginning.
'''

import os
import sys
import win32com.shell.shell as shell
import subprocess
import win32console
import win32event
import win32api
import win32security

SEE_MASK_NOCLOSEPROCESS = 0x00000040
ELEVATE_MARKER = 'win32elevate_marker_parameter'

def areAdminRightsElevated():
    '''
    Tells you whether current script already has Administrative rights.
    '''
    pid = win32api.GetCurrentProcess()
    processToken = win32security.OpenProcessToken(pid, win32security.TOKEN_READ)
    elevated = win32security.GetTokenInformation(processToken, win32security.TokenElevation)
    return bool(elevated)

def waitAndCloseHandle(processHandle):
    '''
    Waits till spawned process finishes and closes the handle for it
    '''
    win32event.WaitForSingleObject(processHandle, win32event.INFINITE)
    processHandle.close()

def elevateAdminRights(waitAndClose=True, reattachConsole=True):
    '''
    This will re-run current Python script requesting to elevate administrative rights.
    
    If waitAndClose is True the process that called elevateAdminRights() will wait till elevated
    process exits and then will quit.
    If waitAndClose is False this function returns None for elevated process and process handle
    for parent process (like POSIX os.fork).
    
    If reattachConsole is False console of elevated process won't be attached to parent process
    so you won't see any output of it.
    '''
    if not areAdminRightsElevated():
        # this is host process that doesn't have administrative rights
        params = subprocess.list2cmdline([os.path.abspath(sys.argv[0])] + sys.argv[1:] + \
                                         [ELEVATE_MARKER])
        res = shell.ShellExecuteEx(fMask=SEE_MASK_NOCLOSEPROCESS, lpVerb='runas', 
                                   lpFile=sys.executable, lpParameters=params)
        if waitAndClose:
            waitAndCloseHandle(res['hProcess'])
            sys.exit(0)
        else:
            return res['hProcess']
    else:
        # This is elevated process, either it is launched by host process or user manually
        # elevated the rights for this script. We check it by examining last parameter
        if sys.argv[-1] == ELEVATE_MARKER:
            # this is script-elevated process, remove the marker
            del sys.argv[-1]
            if reattachConsole:
                # now attach our elevated console to parent's console
                win32console.FreeConsole() # first we free our own (if it exists) console
                win32console.AttachConsole(-1) # then we attach to parent process console

        # indicate we're already running with administrative rights, see docstring
        return None