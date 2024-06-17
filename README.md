# Import and Export Excel in ASP.NET Core 8.0 Razor Pages 📊🚀

NPOI is a powerful and free tool that supports xls, xlsx, and docx extensions. This project is the .NET version of the POI Java project at [Apache POI](http://poi.apache.org/). POI is an open-source project that enables you to read/write XLS, DOC, and PPT files. It covers most features of Excel, such as styling, formatting, data formulas, and extracting images. The best part is, it doesn't require Microsoft Office to be installed on the server. 🌟

With NPOI, you can:

- Generate Excel reports without needing Microsoft Office installed on your server, offering more efficiency than using Microsoft Excel ActiveX in the background.
- Extract text from Office documents to implement full-text indexing features, commonly used to create search engines.
- Extract images from Office documents.
- Generate Excel sheets with formulas.

Let's create an ASP.NET Core Razor Page application. Open Visual Studio 2022, go to File -> New Project -> Select ASP.NET Core Web Application. Enter the project name and click “OK”. Select “Web Application” from the next dialog.

Once the project is created, install the NPOI package for .NET Core by running the following command in the Package Manager Console:

```bash
PM> Install-Package DotNetCore.NPOI
```

### Features 🎉

#### Import Excel Files 📥
We’ll upload an Excel file to the server and then process it using NPOI.

#### Export Excel Files 📤
We’ll create an Excel file with some dummy data using NPOI and download it via the browser.

## Export xlsx/xls in ASP.NET Core

![Export xlsx/xls in ASP.NET Core](/pic/Export-excel-in-ASP.NET-Core-2.0-Razor-Pages.gif)

-------------

## Import xlsx/xls in ASP.NET Core

![Import xlsx/xls in ASP.NET Core](/pic/Import-and-Export-excel-in-ASP.NET-Core-2.0-Razor-Pages.gif)

Happy coding! 💻✨