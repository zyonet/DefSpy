# DefSpy ![Image of Icon](./Resources/DefILSpy-Icon.PNG)
[![Build Status](https://zyonet.visualstudio.com/DefSpy/_apis/build/status/DefSpy-Azure-CI)](https://zyonet.visualstudio.com/DefSpy/_build/latest?definitionId=2)

## Overview

A Visual Studio extension that shows definition of an external type/method under cursor in ILSpy.

The author wants to thank the authors of the following tools for their inspirations:

- GoToILSpy for VS2015: https://marketplace.visualstudio.com/items?itemName=MarekPokornyOVA.GoToILSpyforVS2015
- LinqPad: https://www.linqpad.net/

## How to Use

1. Download and Install DefSpy.vsix

2. Set up path to ILSpy.exe

![Image of Set ILSpy Path](./Resources/SetILSpyPath.PNG)
![Image of File Selection](./Resources/Select-ILSpy-Dlg.PNG)

3. Compile your project, right click a class or method, select "Show Definition in ILSpy..."

![Image of Context Menu](./Resources/right-click-menu.PNG)

4. Decompiled code will be shown in ILSpy

![Image of ILSpy Window](./Resources/ILSpy-Definition-Window.PNG)

5. Happy Coding ;D
