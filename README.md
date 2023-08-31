# Propsheet Demo

## twinBASIC Property Sheet Demo

![image](https://github.com/fafalone/PropsheetDemo/assets/7834493/00ec5873-66ba-40d2-a620-cd013926a01f)

**Requirements:** A recent version of [twinBASIC](https://github.com/twinbasic/twinbasic/releases) and tbShellLib v5.2.210 or higher for all the API definitions-- it's included in the .twinproj file here, but you'll need to add it in a new project. (You can also copy all the relevant defs). Windows XP or newer. 

### Designing the Property Sheet dialog resources
This is a demonstration of using dialog resources with the `PropSheetW` API.\
Since twinBASIC has an excellent Form designer for the regular forms tB/VBx users use 99% of the time, there's no built-in dialog designer and no way to convert a form to a dialog resource. So the first step is preparing the resources using another tool. Visual Studio is a good choice, if a little complicated. You basically need to make a new Windows desktop C++ project, use the designer, then compile and recover the .res file from the compiler intermediates outputs. The advantage of this method is you can assign string IDs like IDC_COMMAND1 and it will autogenerate a resource.h file with `#define` statements, which you can copy/paste into tB and do a couple replace operations to make into Consts. Another option, and the one I used for this project, is the excellent tool [ResourceHacker](http://www.angusj.com/resourcehacker/). You can start with an existing .res file, or create a new one.\
For Visual Studio, there's Property Page templates for the dialog editor; start with one of those, from the Resource View tab in the Solution Explorer, where you can right click the .rc and choose 'Add Resource'. It's then got a familiar enough designer you can figure it out from there.\
To create a new resource in ResourceHacker, use File->New Blank Script, then Action->Add using script template, then choose the DIALOG template. You can then edit it as text, and/or click the green arrow 'Compile' button to convert the .rc script into a .res, which will bring up the visual editor. Right-click the preview to see options to change properties or insert controls, you can move and size controls in the visual editor. The styles for the dialog you want are `DS_SETFONT | WS_CHILD | WS_VISIBLE | WS_CAPTION | WS_SYSMENU`. Once you've first compiled, you can add image resources from the Action menu, then set the caption to the ID when you insert an ICON or BITMAP control. Make sure to set the IDs for each control and image resource, you'll need those to interact with the page.

> [!IMPORTANT]  
> It's important to consider DPI awareness while designing your Property Sheets. If Visual Studio or Resource Hacker are running as DPI aware, and both are by default, your pages will look very different if your application is not DPI aware and your program runs on a zoomed display. Bitmap/icon static controls will be scaled differently, and so will custom font sizes, along with some other things, making your layout have parts cut off as they appear much larger. You can run Resource Hacker with DPI awareness off by going to the exe properties in Explorer, then the Compatibility tab, Change High Dpi Settings, then choose the option to override it's behavior and set to 'System'. 

> [!NOTE]
> You can't set the font of controls individually from the resource file, only the default font for the whole dialog and it's controls. The demo project shows you how to set a different font for an individual control at runtime.

After you have your finished resource file, it's time to import it it into twinBASIC. If you've checked out my [UI Ribbon project](https://github.com/fafalone/UIRibbonDemos), you're already familiar with this interim technique until tB can import .res files directly: Use the template .vbp file from this project, edit it in Notepad to change the name of the resource file, then create new tB project via 'Import from VBP'. Note that this will create a Standard EXE by default, but you can change that to a Standard DLL later if needed, as we'll later do to turn this into a Control Panel Applet in a followup demo. 

Once you've imported the resource file, it's time to start coding! You might want to familiarize yourself with the [Property Sheet Documentation](https://learn.microsoft.com/en-us/windows/win32/controls/property-sheets) on MSDN (I'm not calling it "Learn"). 

### Setting up the code
The first thing to do is add your control IDs. If you used Visual Studio, you can grab them from resouce.h, but from Resource Hacker you'll need to manually define constants for each one. The standard practice is to prefix them with IDC_ for controls, IDI_ for icons, and IDB_ for bitmaps. Next thing to do is set up the code to call the `PropertySheetW` function. You'll need a [`PROPSHEETHEADER`](https://learn.microsoft.com/en-us/windows/win32/controls/pss-propsheetheader) type and an array (if you have more than 1) of [`PROPSHEETPAGE`](https://learn.microsoft.com/en-us/windows/win32/controls/pss-propsheetpage) types. See the links for full documentation of all the possible options. Here's how it's done in the demo:

```
        Dim tPSP(1) As PROPSHEETPAGEW
        Dim tPSH As PROPSHEETHEADERW
    
        'Configure pages
        tPSP(0).dwSize = LenB(Of PROPSHEETPAGEW)
        tPSP(0).dwFlags = PSP_USECALLBACK Or PSP_USETITLE
        tPSP(0).hInstance = hInst
        tPSP(0).pResource = IDD_PROPPAGE_1 'StrPtr(IDD_PROPPAGE_1)
        tPSP(0).pszTitle = StrPtr("Page 1")
        tPSP(0).pfnDlgProc = AddressOf PageOneDlgProc
        tPSP(0).pfnCallback = AddressOf PageOneCallbackProc
        If nShowAs >= pdtWizard97 Then
            tPSP(0).dwFlags = tPSP(0).dwFlags Or PSP_USEHEADERTITLE Or PSP_USEHEADERSUBTITLE
            tPSP(0).pszHeaderTitle = StrPtr("Showing you the wizard...")
            tPSP(0).pszHeaderSubTitle = StrPtr("This is the first page.")
            tPSP(0).bmHeader = StrPtr(IDB_TWINBASIC)
        End If

'repeat for page 2
```

...
