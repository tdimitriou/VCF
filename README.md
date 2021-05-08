# Visual Component Framework (VCF)

A COM based Framework built entirely in VB6 which supports MVVM and a subset of XAML.

**The Framework is built on top of [VBRichClient](http://www.vbrichclient.com) (version 5)**

I started building the framework because I needed to separate the UI from the Business Logic of my commercial application (POS). I also wanted to be able to customize the UI easily without the need to recompile the application each time I needed a change.


The framework has been built based on WPF and the MVVM pattern, however, I did not strictly follow the original implementations.


I am occasionally adding features and performing bud fix and maintenance tasks, according to my needs, but I do not have the time to develop the library on a regular basis. Therefore, I am contributing the sources to the public, hoping that it will be usable to somebody.


The amaizing job that has been done by [Rubberduck](https://rubberduckvba.com/) - (https://github.com/rubberduck-vba/Rubberduck) has brought the VB6 (and VBA) IDE to the next level. Moreover, the recent release of the pre-alpha version of [twinBasic](https://www.twinbasic.com) - (https://github.com/WaynePhillipsEA/twinbasic) and the upcoming release of [RAD Basic](https://www.radbasic.dev) - (https://github.com/radbasic) have raised the expectations that the Classic VB community will soon have a real successor to VB6. I strongly believe that the time is right for all VB6 developers to contribute any project that might be useful, in order to build a strong community. 

### Main Features
* Bindings (including Command Bindings) are supported.
* Dependency Properties are supported, but partially implemented at the moment.
* The XAML resources are stored externally (as XML files) and loaded during component initialization.
* There are also several other (non ui) classes which provide extended functionality.


### Usage
*A sample project demonstrating the base functionality can be found under the __SampleApp__ sub directory.*
