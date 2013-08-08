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

This script performs cleanup of Windows Installer cache trying to be as safe as possible:
it removes only *.msi/*.msp files that are not references as installed on the system (most
likely some leftover junk after unsuccessful installations).

If you break your Windows Installer cache here's a link to MS blog describing the way to fix it:
http://blogs.msdn.com/heaths/archive/2006/11/30/rebuilding-the-installer-cache.aspx
'''
from msi_helpers import getAllPatches, getAllProducts
from win32elevate import elevateAdminRights
from common_helpers import MB
import os
import glob

def getCachedMsiFiles(ext):
    '''
    Finds all cached MSI files at %SystemRoot%\Installer\*.<ext>
    ext can be 'msi' (for installation) or 'msp' (for patches)
    '''
    return [fn.lower() for fn in glob.glob(os.path.join(os.getenv('SystemRoot'), 'Installer',
                                                        '*.%s' % ext))]

def _rotateString(s):
    return ''.join(reversed([''.join(x) for x in zip(*[iter(s)]*2)]))

def unsquishGuid(guid):
    '''
    Unsquishes a GUID (squished GUIDs are used in %SystemRoot%\Installer\$PatchCache$\*
    '''
    squeezedGuid = ''.join(c2 + c1 for (c1, c2) in zip(*[iter(guid)]*2))
    return '{%s}' % '-'.join([_rotateString(squeezedGuid[:8]),
                              _rotateString(squeezedGuid[8:12]),
                              _rotateString(squeezedGuid[12:16]),
                              squeezedGuid[16:20], squeezedGuid[20:]])

def orphanCleanup(name, ext, enumerator):
    files = set()
    for info in enumerator():
        files.add(info.LocalPackage.lower())
    orphanFiles, orphanSize = [], 0
    for fn in getCachedMsiFiles(ext):
        if fn not in files:
            orphanFiles.append(fn)
            orphanSize += os.path.getsize(fn)
    if orphanFiles:
        answer = raw_input('Orphan %s (%d) found occupying %s space. Delete? [y(es)/n(o)] ' % \
                           (name, len(orphanFiles), MB(orphanSize))).lower()
        if answer in ('y', 'yes'):
            for orphan in orphanFiles:
                try:
                    os.remove(orphan)
                except Exception, e:
                    print 'Cannot remove "%s": %r' % (orphan, e)
        else:
            print 'Cancelled by user'
    else:
        print 'Orphan %s not found' % name

def main():
    elevateAdminRights()

    orphanCleanup('patches', 'msp', getAllPatches)
    orphanCleanup('installs', 'msi', getAllProducts)

if __name__ == '__main__':
    main()
