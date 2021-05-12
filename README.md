# Visual Component Framework (VCF)

A COM based Framework built entirely in VB6 which supports the MVVM pattern and a subset of XAML.

**The Framework is built on top of [VBRichClient](http://www.vbrichclient.com) version 5, and an upgrade to version 6 is planned.**

*Code from [vbRichClient/vbWidgets](https://github.com/vbRichClient/vbWidgets) has been used for the implementation of some of the User Controls.*


I started building the framework because I needed to separate the UI from the Business Logic of my commercial application (POS). I also wanted to be able to customize the UI easily without the need to recompile the application each time I needed a change.


The framework has been built based on WPF and the MVVM pattern, however, I did not strictly follow the original implementations. A few workarounds were also necessary in order to overcome some limitations of the VB6 language. 


The project is not yet feature complete, but it is already capable of declaring the whole application UI and bindings in XML for relatively simple user interfaces. The application object can also be defined in XML. This enables the declaration of the Startup Object and several properties and shared (application wide) resources. 

I am occasionally adding features and performing bug fix and maintenance tasks, according to my needs, but I do not have the time to develop the library on a regular basis. Therefore, I am contributing the sources to the public, hoping that it will be useful for direct use or for future development. 


The amaizing job that has been done by [Rubberduck](https://rubberduckvba.com/) - (https://github.com/rubberduck-vba/Rubberduck) has brought the VB6 (and VBA) IDE to the next level. Moreover, the recent release of the pre-alpha version of [twinBasic](https://www.twinbasic.com) - (https://github.com/WaynePhillipsEA/twinbasic) and the upcoming release of [RAD Basic](https://www.radbasic.dev) - (https://github.com/radbasic) have raised the expectations that the Classic VB community will soon have a real successor to VB6. I strongly believe that the time is right for all VB6 developers to contribute any project that might be useful, to start building a strong open source community. 

### Current UI Element Classes
* Window (both top level and child)
* Panel (abstract), Canvas 
* Border
* Button
* Textbox
* Textblock
* Image
* ListView (both bound and unbound)
* UniformGrid
* UserControl
* WindowsFormsHost (to host ActiveX control or any hWnd based window)


### Main Features
* Bindings, including Command Bindings - ICommand Interface and Item Templates for List derived controls are supported.
* Dependency Properties are supported, but partially implemented at the moment.
* The XAML Parser supports extensions through the IMarkupExtension interface. Bindings and Static Resources have been implemented as built in extensions. 
* The XAML resources are stored externally (as XML files) and loaded during component initialization.
* There are also several other non ui classes which provide extended functionality and easy access to Windows API calls. 


### Usage
*A sample project demonstrating the basic functionality can be found under the __SampleApp__ sub directory.*

