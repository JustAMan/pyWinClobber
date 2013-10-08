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
import sys
import errno

class PnpUtilOutputError(Exception):
    pass

class DriverInfo(object):
    '''
    Object that holds information about OEM driver as provided by pnputil.exe -e
    '''
    PARAMS_ORDER = ('name', 'provider', 'driverClass', 'driverDateAndVersion', 'signedBy', None)
    def __init__(self):
        self.name = ''
        self.provider = ''
        self.driverClass = ''
        self.driverDateAndVersion = ''
        self.signedBy = ''
        self.driverDate = None
        self.rawDriverDate = None
        self.driverVersion = ()
        self.__nextParam = 0

    def parseLine(self, line):
        try:
            paramName = self.PARAMS_ORDER[self.__nextParam]
            if not paramName:
                raise Exception()
            setattr(self, paramName, re.search(r'^[^:]*:\s*([^\s\n]?.*)$', line).group(1))
        except:
            raise PnpUtilOutputError(('Bad driver parameters order in line: %s\n' + \
                                      'Tried reading the driver %s') % (line, self))
        else:
            self.__nextParam += 1
        if self.driverDateAndVersion:
            date, version = self.driverDateAndVersion.split(None, 1)
            self.rawDriverDate = date
            self.driverVersion = tuple(int(x) for x in re.findall(r'(\d+)', version))

    def __repr__(self):
        return 'DriverInfo(name=%s, provider=%s, class=%s, version=%s, signed=%s)' % \
                (self.name, self.provider, self.driverClass, self.driverDateAndVersion,
                 self.signedBy)

    def __str__(self):
        if self.driverDate and self.driverVersion:
            date, version = self.driverDate, self.driverVersion
        elif self.driverDateAndVersion:
            date, version = self.driverDateAndVersion.split(None, 1)
        else:
            date, version = '', ''
        return '"%s" by "%s" v%s at %s [%s]' % (self.driverClass, self.provider, version, date,
                                                self.name)

def executePnputil(params):
    '''
    Executes pnputil.exe with given parameters in unicode console, decodes the result in
    Unicode string
    '''
    output = subprocess.check_output(['pnputil'] + params)
    result = []
    brokenLine = False
    # disabling pylint check as it's a false positive
    for line in output.splitlines(): #pylint: disable=E1103
        line = line.strip()
        if not line:
            brokenLine = False
            result.append('')
            continue
        if brokenLine and result:
            result[-1] = result[-1] + line
            brokenLine = False
            continue
        if re.search(r'^[^:]*:\s*$', line):
            # this is broken line, we need to join it with next one
            brokenLine = True
        result.append(line)
    return result

def getAllDrivers():
    '''
    Queries pnputil about all known staged OEM drivers in the system.
    Returns a dictionary that maps oem###.inf file name to DriverInfo() object.
    '''
    try:
        output = executePnputil(['-e'])
    except WindowsError:
        sys.stderr.write('pnputil.exe not found, are you running cleanup of right bitness for '
                         'your system? You need to run 64-bit app on 64-bit system')
        sys.exit(1)
    except subprocess.CalledProcessError, err:
        sys.stderr.write(u'Error calling pnputil.exe: rc = %s, output: %s' % \
                         (err.returncode, err.output))
        sys.exit(1)
    
    if not (' pnp ' in output[0].lower() or 'PnP' in output[0]):
        raise PnpUtilOutputError('Unexpected pnputil.exe output start: %s' % output[0])
    drivers, lastDriver = [], None
    for line in output[1:] + ['']:
        if not line:
            if lastDriver:
                drivers.append(lastDriver)
                lastDriver = None
            continue
        if not lastDriver:
            lastDriver = DriverInfo()
        lastDriver.parseLine(line)

    # now try to guess correct day/month order
    for dateTemplate in ('%d/%m/%Y', '%m/%d/%Y'):
        for driver in drivers:
            try:
                driver.driverDate = datetime.datetime.strptime(driver.rawDriverDate,
                                                               dateTemplate)
            except ValueError:
                break
        else:
            # we didn't encounter any errors while converting data, so we assume that
            # this dateTemplate is the right one, so we stop searching for correct date template
            break
    else:
        # we didn't find suitable date template, notify the user
        raise PnpUtilOutputError('Cannot find suitable date format')
    return {driver.name: driver for driver in drivers}

def deleteDriver(name):
    '''
    Removes staged driver in a safe way, i.e. not forces removal of the driver that is used for
    currently installed devices.
    '''
    print 'Deleting %s...' % name,
    try:
        executePnputil(['-d', name])
    except subprocess.CalledProcessError, err:
        if err.returncode == -536870339:
            # constant retrieved by testing on my machine
            print 'fail: staged driver probably in use'
        else:
            print 'fail: unexpected pnputil return code = %s' % err.returncode
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
    # driver provider (Microsoft, nVidia, etc.) and signed information (MS Compatibility, etc.)
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
        try:
            with open(infName, 'rb') as f:
                content = f.read()
        except IOError, err:
            print 'Warning! Cannot read "%s" file: %s' % (infName, err)
            continue
        infName = os.path.basename(infName)
        try:
            oemFiles[content]
        except KeyError:
            oemFiles[content] = infName
        else:
            # There're two or more exact copies of .inf file with different names, that's really
            # strange. My guess here was that something is wrong with Windows installation,
            # so I used to stop script execution, but for now I've decided to ignore such
            # drivers completely
            continue
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
        try:
            with open(os.path.join(driverRepo, driverDir, infName), 'rb') as f:
                content = f.read()
        except IOError, err:
            if err.errno != errno.ENOENT:
                raise
            # file is missing, skip it
            continue
        try:
            oemName = oemFiles[content]
        except KeyError:
            # this infName is not OEM, skipping
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
                ' (probably superseded by %s)' % oemDups[oemName] if oemName in oemDups else '')

    if dups:
        answer = raw_input(('Possible obsolete drivers found (taking %s). Try to delete? ' + \
                           '[y(es)/n(o)] ') % (MB(dupSize))).lower()
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
