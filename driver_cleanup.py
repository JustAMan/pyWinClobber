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

This script performs DriverStorage cleanup removing possible staged driver duplicates.
The operation should be safe as it utilizes MS pnputil.exe to do the job, and the mode
in which the util is used forbids removing the driver that is currently used for installed
devices.

For more information see "pnputil.exe -?"
'''

from win32elevate import elevateAdminRights
from common_helpers import MB
import subprocess
import re
import os
import glob
import collections
import datetime

class PnpUtilOutputError(Exception):
    pass

class DriverInfo(object):
    '''
    Object that holds information about OEM driver as provided by pnputil.exe -e
    '''
    PARAMS_ORDER = [('name', 'Published name'),
                    ('provider', 'Driver package provider'),
                    ('driverClass', 'Class'),
                    ('driverDateAndVersion', 'Driver date and version'),
                    ('signedBy', 'Signer name'),
                    (None, '')]
    def __init__(self):
        self.name = ''
        self.provider = ''
        self.driverClass = ''
        self.driverDateAndVersion = ''
        self.signedBy = ''
        self.driverDate = None
        self.driverVersion = ()
        self.__nextParam = 0

    def parseLine(self, line):
        try:
            paramName, paramText = self.PARAMS_ORDER[self.__nextParam]
            if not paramName:
                raise Exception()
            setattr(self, paramName,
                    re.search(r'^{0}\s*:\s*([^\s\n]?.*)$'.format(re.escape(paramText)),
                              line).group(1)
                    )
        except:
            raise PnpUtilOutputError(('Bad driver parameters order in line: %s\n' + \
                                      'Tried reading the driver %s') % (line, self))
        else:
            self.__nextParam += 1
        if self.driverDateAndVersion:
            date, version = self.driverDateAndVersion.split()
            self.driverDate = datetime.datetime.strptime(date, '%m/%d/%Y')
            self.driverVersion = tuple(int(x) for x in re.findall(r'(\d+)', version))

    def __repr__(self):
        return 'DriverInfo(name=%s, provider=%s, class=%s, version=%s, signed=%s)' % \
                (self.name, self.provider, self.driverClass, self.driverDateAndVersion,
                 self.signedBy)

    def __str__(self):
        date, version = self.driverDateAndVersion.split()
        return '"%s" by "%s" v%s at %s [%s]' % (self.driverClass, self.provider, version, date,
                                                self.name)

def getAllDrivers():
    '''
    Queries pnputil about all known staged OEM drivers in the system.
    Returns a dictionary that maps oem###.inf file name to DriverInfo() object.
    '''
    output = subprocess.check_output(['pnputil', '-e']).splitlines()
    if output[0].strip() != 'Microsoft PnP Utility':
        raise PnpUtilOutputError('Unexpected pnputil.exe output start: %s' % output[0])
    drivers, lastDriver = [], None
    for line in output[1:] + ['']:
        if not line.strip():
            if lastDriver:
                drivers.append(lastDriver)
                lastDriver = None
            continue
        if not lastDriver:
            lastDriver = DriverInfo()
        lastDriver.parseLine(line)
    return {driver.name: driver for driver in drivers}

def deleteDriver(name):
    '''
    Removes staged driver in a safe way, i.e. not forces removal of the driver that is used for
    currently installed devices.
    '''
    print 'Deleting %s...' % name,
    try:
        subprocess.check_output(['pnputil', '-d', name])
    except subprocess.CalledProcessError, err:
        if 'One or more devices are presently installed using the specified INF' in err.output:
            print 'fail: staged driver probably in use'
        else:
            print 'fail: unexpected pnputil return code = %s, output:' % err.returncode
            print err.output
        return False
    else:
        print 'done'
        return True

def getFolderSize(path):
    '''
    Calculates target path size (recursively if target is a directory)
    '''
    result = os.path.getsize(path)
    if os.path.isdir(path):
        for root, dirs, files in os.walk(path):
            for node in files + dirs:
                result += os.path.getsize(os.path.join(root, node))
    return result

def main():
    '''
    Main function for the script
    '''
    elevateAdminRights()

    print 'Reading all OEM drivers...',
    drivers = getAllDrivers()
    print 'done'
    
    # Let's find possible duplicates. The tuple of driver class (e.g. Keyboard, Display, etc.),
    # driver provider (Microsoft, Nvidia, etc.) and signed information (MS Compatibility, etc.)
    # is considered to be the key defining a driver for the device. All drivers that have this
    # key being the same are considered to be the instances of the same driver, thus we sort
    # them by version and date and mark all older ones as duplicates of the most recent driver.
    duplicates = collections.defaultdict(list)
    for driver in drivers.itervalues():
        duplicates[(driver.driverClass, driver.provider, driver.signedBy)].append(driver)
    oemDups = {}
    for key, driversList in duplicates.items():
        if len(driversList) <= 1:
            del duplicates[key]
        else:
            driversList.sort(cmp=lambda d1, d2: cmp(d1.driverVersion, d2.driverVersion) or \
                                                cmp(d1.driverDate, d2.driverDate),
                             reverse=True)
            for dupDriver in driversList[1:]:
                oemDups[dupDriver.name] = driversList[0].name

    # Now we read all %SystemRoot%\inf\oem*.inf files to make a map that will allow us by
    # estimating the size of drivers stored in DriverStore to find out which oem drivers are
    # the largest and what we should remove.
    print 'Reading oem*.inf files...',
    infFiles = os.path.join(os.getenv('SystemRoot'), 'inf', 'oem*.inf')
    oemFiles = {}
    for infName in glob.glob(infFiles):
        with open(infName, 'rb') as f:
            content = f.read()
        infName = os.path.basename(infName)
        try:
            dupFile = oemFiles[content]
        except KeyError:
            oemFiles[content] = infName
        else:
            # There're two exact copies of .inf file with different names, that's really
            # strange. Our guess here is that something is wrong with Windows installation,
            # so we stop our execution
            raise Exception('%s is duplicate of %s' % (infName, dupFile))
    print 'done'
    
    # now parse %SystemRoot%\system32\DriverStore\FileRepository
    print 'Parsing DriverStore...',
    driverRepo = os.path.join(os.getenv('SystemRoot'), 'system32', 'DriverStore',
                              'FileRepository')
    driverSize = []
    for driverDir in os.walk(driverRepo).next()[1]:
        # All folders should in here should have the same pattern - abc.inf_something where
        # abc.inf lies within and should match to some oem###.inf file read above if this driver
        # is OEM (not built in current Windows setup).
        try:
            infName = re.match(r'^(.*?\.inf)_.*$', driverDir).group(1)
        except ValueError:
            # this folder does not match desired pattern, ignore it
            continue
        with open(os.path.join(driverRepo, driverDir, infName), 'rb') as f:
            content = f.read()
        try:
            oemName = oemFiles[content]
        except KeyError:
            # this infName is not oem, skipping
            continue
        driverSize.append((oemName, getFolderSize(os.path.join(driverRepo, driverDir))))
    print 'done'
    
    print 'Drivers (sorted by size):'
    driverSize.sort(reverse=True, key=lambda (oemName, size): size)
    dups, dupSize = [], 0
    for oemName, size in driverSize:
        if oemName in oemDups:
            dups.append((oemName, size))
            dupSize += size
        print '%s: %s%s' % (drivers[oemName], MB(size),
                ' (duplicate of %s)' % oemDups[oemName] if oemName in oemDups else '')

    if dups:
        answer = raw_input('Possible duplicates found (taking %s). Delete? [y(es)/n(o)] ' % \
                           (MB(dupSize))).lower()
        cleanedSize = 0
        if answer in ('y', 'yes'):
            for dup, size in dups:
                if deleteDriver(dup):
                    cleanedSize += size
            print 'Was able to clean up %s out of %s expected' % (MB(cleanedSize), MB(dupSize))
        else:
            print 'Cancelled by user'

if __name__ == '__main__':
    main()
