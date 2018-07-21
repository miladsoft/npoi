Import and Export excel in ASP.NET Core 2.0 Razor Pages
===========
NPOI is a free tool which supports xls, xlsx and docx extensions. This project is the .NET version of POI Java project at http://poi.apache.org/. POI is an open source project, which can help you read/write XLS, DOC, PPT files. It covers most features of Excel like styling, formatting, data formulas, extract images, etc.. The good thing is, it doesn’t require Microsoft Office to be installed on the server.

For example, you can use it to Generate an Excel report without Microsoft Office suite installed on your Server and more efficient than calling Microsoft Excel ActiveX in the background.
    Extract text from Office documents to help you implement full-text indexing feature (most of the time this feature is used to create search engines).
    Extract images from Office documents.
    Generate Excel sheets, which contains formulas.

Let’s create an ASP.NET Core Razor Page application. Open Visual Studio 2017. Hit File-> New Project -> Select ASP.NET Core Web Application. Enter the project name and say “OK”. Select “Web Application” from the next dialog. 

Once the project is created, install NPOI package for .NET Core. To install, run the following command in the Package Manager Console:

```
PM> Install-Package DotNetCore.NPOI
```

To show,

- Import: We’ll upload an excel file on the server and then process it using NPOI.
- Export: We’ll create an excel file with some dummy data using NPOI and download the same in the browser.


# Export xlsx/xls in ASP.NET Core
![Export xlsx/xls in ASP.NET Core](/pic/Export-excel-in-ASP.NET-Core-2.0-Razor-Pages.gif)

-------------

# Import xlsx/xls in ASP.NET Core
![Import xlsx/xls in ASP.NET Core](/pic/Import-and-Export-excel-in-ASP.NET-Core-2.0-Razor-Pages.gif)
