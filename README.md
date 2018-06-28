# PageScriptControl
control your web page with javascript

embed brower into WPF window, expose HTML function and other function to javascript engine, and then you can control the web page with javascript. Now we have add following function to javascript engine, you can also add more function

log(message) log message to console

doc Html document object

go(url) browser navigate function

value(selector, value) set value to html element

click(selector) click button or html element

ready() wait for browser load finish

view(selector) scroll html element to view

# Demo
if you want to automatically login to www.newsmth.com, you can use following script
```csharp
go("http://www.newsmth.com");
ready();
value("#id","user name");
value("#pwd","password");
click("#b_login");
ready();
log("login");
```
