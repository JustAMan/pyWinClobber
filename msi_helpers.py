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

Helper module that allows to enumerate MSI patches and installs and access their properties
using MSI API
'''

import ctypes
from ctypes.wintypes import DWORD
from ctypes import c_char_p, POINTER, c_uint, pointer
from win32elevate import elevateAdminRights

LPDWORD = POINTER(DWORD)

MsiEnumPatchesEx = ctypes.windll.msi.MsiEnumPatchesExA
MsiEnumPatchesEx.argtypes = (c_char_p, c_char_p, DWORD, DWORD, DWORD, c_char_p, c_char_p,
                             LPDWORD, c_char_p, LPDWORD)
MsiEnumPatchesEx.restype = c_uint

MsiGetPatchInfoEx = ctypes.windll.msi.MsiGetPatchInfoExA
MsiGetPatchInfoEx.argtypes = (c_char_p, c_char_p, c_char_p, DWORD, c_char_p, c_char_p, LPDWORD)
MsiGetPatchInfoEx.restype = c_uint

# from MSDN
ALL_USERS = 's-1-1-0'
MSIINSTALLCONTEXT_ALL = 1 | 2 | 4
MSIINSTALLCONTEXT_MACHINE = 4
MSIPATCHSTATE_ALL = 15

ERROR_NO_MORE_ITEMS = 0x103

class PatchInfo(object):
    def __init__(self, patchGuid, productGuid, dwContext, userSid):
        self.__patchGuid = patchGuid
        self.__productGuid = productGuid
        self.__userSid = userSid
        self.__dwContext = dwContext

    def __str__(self):
        return 'Patch: %s, product: %s (by %s)' % (self.__patchGuid, self.__productGuid,
                                                   self.__userSid or '<system>')
    
    def __getattr__(self, name):
        buffSize = DWORD(10)
        userSid = self.__userSid if self.__dwContext != MSIINSTALLCONTEXT_MACHINE else None
        result = MsiGetPatchInfoEx(self.__patchGuid, self.__productGuid, userSid,
                                   self.__dwContext, str(name), None, pointer(buffSize))
        if result != 0 or buffSize.value == 0:
            raise AttributeError('%s is missing %s (error: %s)' % (self, name, result))
        buffSize = DWORD(buffSize.value + 1)
        buff = ctypes.create_string_buffer(buffSize.value)
        result = MsiGetPatchInfoEx(self.__patchGuid, self.__productGuid, userSid,
                                   self.__dwContext, str(name), buff, pointer(buffSize))
        if result != 0:
            raise AttributeError('Cannot get %s property for %s: error %s' % \
                                 (name, self, result))
        return buff.value
    
def allPatches():
    index = 0
    # Allocate big enough buffer to keep GUID plus null terminator
    patchGuid = ctypes.create_string_buffer(len('{01234567-89AB-CDEF-0123-456789ABCDEF}') + 1)
    productGuid = ctypes.create_string_buffer(len('{01234567-89AB-CDEF-0123-456789ABCDEF}') + 1)
    userSidSize = DWORD(10)
    dwContext = DWORD(111)
    while True:
        result = MsiEnumPatchesEx(None, ALL_USERS, MSIINSTALLCONTEXT_ALL, MSIPATCHSTATE_ALL,
                                  index, patchGuid, productGuid, ctypes.byref(dwContext), None,
                                  ctypes.byref(userSidSize))
        if result != 0:
            if result != ERROR_NO_MORE_ITEMS:
                raise Exception('MsiEnumPatchesEx unexpectedly returned %s' % result)
            break

        if userSidSize.value != 0:
            userSidSize = DWORD(userSidSize.value + 1)
            userSid = ctypes.create_string_buffer(userSidSize.value)
            result = MsiEnumPatchesEx(None, ALL_USERS, MSIINSTALLCONTEXT_ALL, MSIPATCHSTATE_ALL,
                                      index, patchGuid, productGuid, None, userSid,
                                      ctypes.byref(userSidSize))
        if result == 0:
            index += 1
            yield PatchInfo(patchGuid.value, productGuid.value, dwContext.value,
                            userSid.value if userSidSize else '')
        else:
            raise Exception('Cannot get needed szTargetUserSid size: error = %s' % result)
        
if __name__ == '__main__':
    elevateAdminRights()
    for patchInfo in allPatches():
        print '%s: package = %s' % (str(patchInfo), patchInfo.LocalPackage)
    