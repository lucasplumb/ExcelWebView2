# ExcelWebView2
An embedded WebView2 browser project for Microsoft Excel.

![examples](https://github.com/lucasplumb/ExcelWebView2/assets/8316592/cac706d0-8d15-4139-9f7d-d294cc3f6138)

# What is this?

ExcelWebView2 is intended to be a framework for creating embedded browser projects using Edge/WebView2 within Microsoft Excel.

The goal is to support automated browser tasks using VBA, which may have previously used the embedded Internet Explorer control, but for modern websites that no longer support IE.

This is a task which has not seen any real support from Microsoft - so until now the only real option has been Selenium - unfortunately, Selenium uses CDP which may not be enabled for all users, and only provides interaction with an external browser, rather than an embedded browser object.

This project was created using 32-bit Office 2016, and WebView2Loader.dll version 1.0.1072.54.  I have included the .dll in the project files for convenience, but you can also download it via the official Microsoft WebView2 Runtime, or scan your computer for an existing version (you probably have one already, though I make no guarantees this project will work with any other version than 1.0.1072.54)



# How do I use it?

The first thing you that should be noted, is that this project is in the very early stages of development.  Not all features of WebView2 have been implemented.  If you receive any sort of "Automation error", most likely it means that the Type Library has not been modified to support that method/property.

That being said, working with this project should mostly be left to more advanced VBA developers with an understanding of COM.  I also provide no guarantee that the code is bug free or thoroughly tested, so please browse the project and create PR's for any issues you see.

## Plugins

I advise developers first explore the provided "plugin" functionality of the project.

New plugins should be based on the "pluginBase" class.  Copy and paste the code into a NEW class module.

![creating_plugin](https://github.com/lucasplumb/ExcelWebView2/assets/8316592/dbd53d60-4b6e-4980-85dd-c9512fe956f1)

I've created an example plugin that demonstrates a few functions to get you started working with the project, called "pluginExampleCls".

Note that I've copied the code from "pluginBase" into a new class module, and changed the "pluginInterface_newInstance()" property as such:

![creating_plugin1](https://github.com/lucasplumb/ExcelWebView2/assets/8316592/07ef8649-88c8-4e6f-8373-fcf7ba620cca)

In order to load my plugin to work with the rest of the code in the project, I need to create an entry for it in the "pluginLoader" module.

Note the following image where I've added my example plugin class to the LoadPlugins() method of the pluginLoader module.

![loading_plugin](https://github.com/lucasplumb/ExcelWebView2/assets/8316592/3d3a3aaf-a832-4700-91d8-41cc2d51cde1)

This new plugin provides extended functionality for handling the events that WebView2 raises.  However, it is unadvisable to write new functions or keep track of any sort of "state" in the plugin class itself without modifications to other code in the project.

For now, we should create a new standard module to perform our automation tasks which is provided information by our plugin class module.  See the pluginExample standard module:

![creating_plugin_example](https://github.com/lucasplumb/ExcelWebView2/assets/8316592/e6eadd13-3550-483d-8c9e-3e4b38b14c23)

## More to come

I hope to provide more detailed explanations and tutorials for the code in the future.  For now, I would just like to get this released and see some feedback.  Thanks and happy coding!
