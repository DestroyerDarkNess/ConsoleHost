# ConsoleHost
ConsoleHost control / Embed Console In Winform Project .
Remember to leave your Star to the Project! Thank you!

## Introduction
ControlHost allows you to enter a console as if it were a Winform control.
By default it loads a basic console. If the load fails, The plugin loads a HOST CMD.

You can also customize the process to be Hosted using the TargetProcess property of the control.

## Preview

![Preview](https://i.ibb.co/3STs3Ks/ss1.png)

![Preview](https://i.ibb.co/pJhbLk8/ss2.png)

## Known issues :

On some computers for some reason it does not write the Console.Writeline method. For this problem, it has been created in Start method:
```vb
ConsoleHost1.Unsecure_Initialize()
```

<hr style="background-color:blue;"></hr>

If the problem persists follow the steps below:

Enabling native code debugging redirects the console to the Output window. However, regardless of the native code debugging settings, I saw absolutely no results in any of the places until I enabled the Visual Studio hosting process.

This could have been the reason why simply disabling native code debugging didn't solve your problem.

  Go to Project Properties -> Debug -> Enable Debuggers and make sure 'Enable Visual Studio Hosting Process' is checked.

Also enabling "Sql Server Debugging" prevents the console from working, make sure it is disabled.

## Possible uses

- Custom CMD For script loading projects, Possibly you can create your own Script language and want to use ConsoleHost.
- Custom development of Debuggers with console design.
- Everything what your imagination wants.

## Contributors
- Destroyer : Creator and Developer.  / Discord : Destroyer#8328 
- PinkAxol : She suggests this function, and I implement it, so it was his Idea or something ... / Discord : PinkAxol#9856

  ## Special thanks :
- [ElektroStudios](https://github.com/ElektroStudios): For its Class SetWindowState and other Functions.
   - Which are part of its [DevCase Framework](https://codecanyon.net/item/elektrokit-class-library-for-net/19260282)
