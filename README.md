XL Toolbox NG (Next Generation)
===============================

The XL Toolbox NG source code is written in (mostly) C# with Visual Studio
Professional 2013 and target for the .NET framework 4.0.

For more information about this project, see <http://xltoolbox.sf.net>.

This project uses the Git source code management system. You can find the
repository at <https://sf.net/p/xltoolbox/ng-code>.


Code-signing the binaries
-------------------------

The sources do of course not include the confidential strong name key (.snk)
file that is needed to sign the binaries. If you want to build the solution
yourself, you have different options:

- Unzip 'source.zip' or clone the Git repository and subsequently remove
  the code signing option from all of the project properties to build
  unsigned binaries. If cloning the Git repository, it is best to use a
  separate branch to make the changes to the projects properties. If you
  later update the repository from remote, you can git-rebase this
  branch on top of HEAD.
- Unzip 'source.zip' in Windows and supply a strong name key file in
  every subdirectory of the extracted source tree. The Visual Studio
  project properties expect the file name to be "xltb.snk". The original
  strong name key file is not included in the distributed sources for
  obvious reasons.
- Clone the Git repository on a \*nix file system. The repository
  contains symbolic links named "xltb.snk" in the subdirectories that
  point to a strong name key file in an unrelated directory
  "../private/" that lies outside of the repository. Therefore you would
  only need to create such directory and put the strong name key file in
  there. For instance:
	  
		# make new directory that holds everything and enter it
		mkdir XLToolbox  
		chdir XLToolbox

		# clone the repository into `source`
		git clone git://git.code.sf.net/p/xltoolbox/ng-code source

		# make directory for strong name key file
		mkdir private 

		# Then, start Windows and create a new strong name key file
		# named `xltb.snk` in the `private` directory.


Creating a strong name key file
-------------------------------

Visual Studio comes with the `sn.exe` tool that you can use to create a .snk
file. It is somewhat hidden; you may for example find it in:

		C:\Program Files (x86)\Microsoft SDKs\Windows\v8.1A\bin	

In the command window, `cd` to this directory and execute:

		sn.exe -k xltb.snk
		move xltb.snk <drive:\path\to\sources_or_private>

Whether you have to move the .snk file to the `private` directory or to the
source directory depends on what method you have chosen above. On
Windows-only systems where you will likely not be using symlinks, you need
to copy the .snk file to each and every subdirectory in the source tree.


Note
----

It should go without saying that you cannot of course mix binaries from the
original distribution with binaries that you build yourself, as they do not
share the same strong name key.


License
-------

    Daniel's XL Toolbox NG
    Copyright (C) 2008-2014  Daniel Kraus  <xltoolbox@gmx.net>

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.

