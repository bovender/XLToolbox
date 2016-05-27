XL Toolbox NG (Next Generation) ![Downloads of the NG series](https://img.shields.io/github/downloads/bovender/xltoolbox/total.svg?maxAge=2592000)
===============================

##### Data analysis and visualization for Excel&reg;, for free.

> For more information about this project, see <https://www.xltoolbox.net>.

The XL Toolbox NG source code is written in (mostly) C# with Visual Studio
Professional 2013 and targeted for the .NET framework 4.0.


Building the Debug configuration
--------------------------------

The Visual Studio project files have been hand-edited to use different
sources of the [Bovender][] assembly: In the _Debug_ configuration, the
assembly is referenced locally. In the _Release_ configuration, the
assembly is referenced as a NuGet package. This has the advantage (for
me) that I can make changes to [Bovender][] without having to release a
new NuGet package every time I want to test it in the NG solution.

This is my directory structure:

    x:\
      Code
        bovender\
          Bovender\
          BovenderUnitTests\
        xltoolbox\
          NG\
            XLToolbox\
            XLToolboxForExcel\
            UnitTest\

The bovender assembly is thus referenced from the XLToolbox project file
(which resides in `x:\Code\xltoolbox\NG\XLToolbox`) as
`..\..\..\bovender\Bovender\bovender.csproj`.

If you want to build with _Debug_ configuration, but do not have the
[Bovender][] assembly in the same directory structure as I have, you can
(and should) use the _Debug (Bovender via NuGet)_ configuration.


Code-signing the binaries
-------------------------

The sources do of course not include the confidential strong name key (.snk)
file that is needed to sign the binaries. If you want to build the solution
yourself, you have different options:

- Clone the Git repository and subsequently remove
  the code signing option from all of the project properties to build
  unsigned binaries. It is best to use a
  separate branch to make the changes to the projects properties. If you
  later update the repository from remote, you can git-rebase this
  branch on top of HEAD.
- On Windows, obtain the sources and supply a strong name key file in
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
    Copyright (C) 2014-2016  Daniel Kraus <bovender@bovender.de>, <daniel@xltoolbox.net>

	Licensed under the Apache License, Version 2.0 (the "License");
	you may not use this file except in compliance with the License.
	You may obtain a copy of the License at

	    http://www.apache.org/licenses/LICENSE-2.0

	Unless required by applicable law or agreed to in writing, software
	distributed under the License is distributed on an "AS IS" BASIS,
	WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
	See the License for the specific language governing permissions and
	limitations under the License.

[Bovender]: https://github.com/bovender/bovender
<!-- vim: set tw=72 ai sw=2 ts=2 : -->
