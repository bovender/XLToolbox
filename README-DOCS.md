\mainpage

This is the documentation of the source code of Daniel's XL Toolbox,
Next Generation (NG).

The solution consists of three main C# projects and two test projects:

- `Bovender` is my assembly with reusable code; most importantly, my
  Model--View--View Model framework is defined in here (see namespace
  Mvvm).
- `XLToolbox` contains almost all the Excel-specific code. A central
  class is the Dispatcher class which provides user-entry points.
- `Addin` is a small assembly that is called by the VSTO engine; it
  implements an ThisAddIn_Startup method and a Ribbon. The Ribbon calls
  into the Dispatcher class defined in the XLToolbox assembly.
- `UnitTests` contains tests for the XLToolbox assembly.
- `BovenderUnitTests` contains tests for the Bovender assembly.

There are two main namespaces:

- XLToolbox
- Bovender

The tests are in the XLToolbox.Test and Bovender.Test namespaces.

When I have time, I will describe my MVVM framework further.

--Daniel

-----

	Copyright 2014-2015 Daniel Kraus

	Licensed under the Apache License, Version 2.0 (the "License");
	you may not use this file except in compliance with the License.
	You may obtain a copy of the License at

		http://www.apache.org/licenses/LICENSE-2.0

	Unless required by applicable law or agreed to in writing, software
	distributed under the License is distributed on an "AS IS" BASIS,
	WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
	See the License for the specific language governing permissions and
	limitations under the License.
<!-- vim: set tw=72 : -->
