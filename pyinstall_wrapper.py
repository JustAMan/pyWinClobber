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

Helper script to build standalone executables using PyInstaller, see
http://www.pyinstaller.org/ for more information about it.
'''

import sys
import subprocess
import os
import re
import shutil

class PyInstallerWrap(object):
    def __init__(self, pyInstallerDir):
        self.dir = pyInstallerDir

    def createSpec(self, script):
        subprocess.check_call([sys.executable, os.path.join(self.dir, 'utils', 'Makespec.py'),
                               script, '-n', 'pyWinClobber'])

    def _parseSpec(self, source):
        with open('%s.spec' % source, 'r') as f:
            data = f.read()
        #return data.replace('a = Analysis', '{0}_a = Analysis').replace('a.', '{0}_a.').\
        #            replace('pyz', '{0}_pyz').replace(
        match = re.search(r'^(?P<start>.*?)(?P<end>pyz = PYZ.*)', data, re.DOTALL).groupdict()
        start = match['start'].replace('a = Analysis', '{0}_a = Analysis')
        end = match['end']
        for analysisAttr in ('pure', 'scripts', 'binaries', 'zipfiles', 'datas'):
            end = end.replace('a.%s' % analysisAttr, '{0}_a.%s' % analysisAttr)
        end = end.replace('pyz = PYZ', '{0}_pyz = PYZ').\
                  replace('exe = EXE(pyz', '{0}_exe = EXE({0}_pyz').\
                  replace('coll = COLLECT(exe', '{0}_coll = COLLECT({0}_exe')
        return start, end

    def mergeSpecs(self, targetSpec, sourceSpecs):
        analysis, finish, merge = [], [], []
        for idx, spec in enumerate(sourceSpecs):
            name = 'a_%d' % idx
            start, end = self._parseSpec(spec)
            analysis.append(start.format(name))
            finish.append(end.format(name))
            merge.append('(%s_a, "%s", "%s")' % (name, spec, spec))
        with open('%s.spec' % targetSpec, 'w') as out:
            out.writelines(analysis)
            out.write('MERGE( %s )\n' % ', '.join(merge))
            out.writelines(finish)

    def buildBundle(self, spec):
        subprocess.check_call([sys.executable, os.path.join(self.dir, 'pyinstaller.py'), spec,
                               '-y'])

    def mergeBinaries(self, sources, target):
        os.makedirs(target)
        for srcDir in sources:
            src = os.path.join('dist', srcDir)
            for fn in os.walk(src).next()[2]:
                shutil.copy2(os.path.join(src, fn), target)

    def prepareWipe(self, target):
        for directory in ('build', 'dist', target):
            if os.path.isdir(directory):
                shutil.rmtree(directory)

SCRIPTS = ('msi_cleanup', 'driver_cleanup')

def main():
    if len(sys.argv) != 2:
        sys.stderr.write('usage: %s path-to-pyInstaller-dir' % sys.argv[0])
        sys.exit(1)

    wrapper = PyInstallerWrap(os.path.abspath(sys.argv[-1]))
    targetDir = os.path.join('release', '%s-bit' % ('32' if sys.maxsize == 2**31 - 1 else '64'))
    wrapper.prepareWipe(targetDir)
    for script in SCRIPTS:
        wrapper.createSpec('%s.py' % script)
    wrapper.mergeSpecs('pyWinClobber', SCRIPTS)
    wrapper.buildBundle('pyWinClobber.spec')
    wrapper.mergeBinaries(SCRIPTS, targetDir)

if __name__ == '__main__':
    main()