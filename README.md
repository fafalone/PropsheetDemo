# Propsheet Demo

## twinBASIC Property Sheet Demo

![image](https://github.com/fafalone/PropsheetDemo/assets/7834493/00ec5873-66ba-40d2-a620-cd013926a01f) ![image](https://github.com/fafalone/PropsheetDemo/assets/7834493/353db63d-6a07-47bc-8de1-64723f64fa1a)


**Requirements:** A recent version of [twinBASIC](https://github.com/twinbasic/twinbasic/releases) and tbShellLib v5.2.210 or higher for all the API definitions-- it's included in the .twinproj file here, but you'll need to add it in a new project. (You can also copy all the relevant defs). Windows XP or newer. 

### Designing the Property Sheet dialog resources
This is a demonstration of using dialog resources with the `PropSheetW` API, which uses Dialog resources rather than traditional windows.\
Since twinBASIC has an excellent Form designer for the regular forms tB/VBx users use 99% of the time, there's no built-in dialog designer and no way to convert a form to a dialog resource. So the first step is preparing the resources using another tool. Visual Studio is a good choice, if a little complicated. You basically need to make a new Windows desktop C++ project, use the designer, then compile and recover the .res file from the compiler intermediates outputs. The advantage of this method is you can assign string IDs like IDC_COMMAND1 and it will autogenerate a resource.h file with `#define` statements, which you can copy/paste into tB and do a couple replace operations to make into Consts. Another option, and the one I used for this project, is the excellent tool [ResourceHacker](http://www.angusj.com/resourcehacker/). You can start with an existing .res file, or create a new one.\
For Visual Studio, there's Property Page templates for the dialog editor; start with one of those, from the Resource View tab in the Solution Explorer, where you can right click the .rc and choose 'Add Resource'. It's then got a familiar enough designer you can figure it out from there.\
To create a new resource in ResourceHacker, use File->New Blank Script, then Action->Add using script template, then choose the DIALOG template. You can then edit it as text, and/or click the green arrow 'Compile' button to convert the .rc script into a .res, which will bring up the visual editor. Right-click the preview to see options to change properties or insert controls, you can move and size controls in the visual editor. Any control that can be created by your app with `CreateWindowEx` can be used, though only the common controls are available in the visual editors. The styles for the dialog you want are `DS_SETFONT | WS_CHILD | WS_VISIBLE | WS_CAPTION | WS_SYSMENU`. Once you've first compiled, you can add image resources from the Action menu, then set the caption to the ID when you insert an ICON or BITMAP control. Make sure to set the IDs for each control and image resource, you'll need those to interact with the page.

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
        tPSP(0).pResource = IDD_PROPPAGE_1
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

        'Configure header        
        tPSH.dwSize = LenB(Of PROPSHEETHEADERW)
        tPSH.dwFlags = PSH_USECALLBACK Or PSH_PROPSHEETPAGE Or PSH_USEICONID
        If nShowAs = pdtModeless Then tPSH.dwFlags = tPSH.dwFlags Or PSH_MODELESS
        If nShowAs = pdtWizard Then tPSH.dwFlags = tPSH.dwFlags Or PSH_WIZARD
        If nShowAs = pdtWizard97 Then
            tPSH.dwFlags = tPSH.dwFlags Or PSH_WIZARD97
            tPSH.hbmHeader = StrPtr(IDB_TWINBASIC)
        End If
        If nShowAs = pdtAeroWizard Then
            tPSH.dwFlags = tPSH.dwFlags Or PSH_AEROWIZARD Or PSH_WIZARD
            tPSH.hbmHeader = StrPtr(IDB_TWINBASIC)
        End If
        tPSH.hInstance = hInst
        tPSH.hIcon = StrPtr(IDI_TWINBASIC)
        tPSH.pszCaption = StrPtr(szTitle)
        tPSH.nPages = UBound(tPSP) + 1
        tPSH.hwndParent = hWnd
        tPSH.pfnCallback = AddressOf PropsheetCallback
        tPSH.ppsp = VarPtr(tPSP(0))
```

You'll definitely need the `pfnDlgProc` functions, but the other callbacks are optional, they're mainly to provide extra information for development and debugging here. **Reminder:** Since tB doesn't yet support unions, I had to picked one or the other for which name to use, and it may not always be right, but it is the same effect: For example, `tPSH.hIcon` is not set to an `HICON`, but to a resource name, so in a union-supporting language, you'd use `tPSH.pszIcon` instead, but since they're at the same offset and the `PSH_USEHICON` flag is not specified, it's `pszIcon` as far as Windows is concerned.

> [!IMPORTANT]
> If you do use callbacks for the pages, make sure to return 1 in response to `PSPCB_CREATE`, otherwise the page will not be created and the call will fail. I had some trouble with this because of an odd situation where `Debug.Print` statements weren't executing in it while they were from other messages, so I had incorrectly assumed the code never even reached that point so never double checked it. Special thanks to The trick on VBForums for figuring that out.

The dialog procs are very similar to subclassing WndProcs:

```
    Private Function PageOneDlgProc(ByVal hDlg As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
        Select Case uMsg
            Case WM_INITDIALOG
                Debug.Print "WM_INITDIALOG on p1"
                hDialog = hDlg
                InitPageOneValues hDlg
```
The `WM_INITDIALOG` message is the best place to set the initial states of all your controls. There's a number of dialog-specific functions for this, allowing you to do things by ID instead of hWnd, but for other things you can get the hWnd of a control with the `GetDlgItem` function. We store the handle to the dialog in a module-level variable for use in the Propsheet callback, because each page has it's own handle (remember, each page is a separate dialog resource), and the PropSheetCallback passes only the top level hwnd. Setting the initial controls states is fairly self explanatory; I'll point out one thing as noted above, changing the font on an individual control, in this case our page one welcome message:

```
        Dim hFontHdr As LongPtr = CreateFontW(-20, 0, 0, 0, FW_BOLD, CFALSE, CTRUE, CFALSE, _
                                                ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, _
                                                CLEARTYPE_QUALITY, DEFAULT_PITCH Or FF_SWISS, _
                                                StrPtr("Segoe UI"))

        If hFontHdr Then
            SendDlgItemMessageW hDlg, IDC_HDRPAGE1, WM_SETFONT, hFontHdr, ByVal CTRUE
            SetDlgItemTextW hDlg, IDC_HDRPAGE1, StrPtr("Welcome to our twinBASIC Property Sheet Demo!")
            'DeleteObject hFontHdr '<-- Don't do this here! If the font is deleted the font reverts
```
There are of course other ways to create a font, including from the StdFont object to match a Form's font, but that's a separate topic. If you right click the API names for `SendDlgItemMessageW` or `SetDlgItemTextW` and click Go to definition, it will take you to the Dialog section of tbShellLib, where you can see all of the available dialog-specific APIs.\
After that, we monitor for changes the exact same we would for an API-created window or a subclassed window, primarily through `WM_COMMAND` and `WM_NOTIFY` messages, depending on the control. We get the current values, compare it to the values we loaded in the InitPageNValues functions if neccessary (i.e. if a click isn't dispositive of whether we want to consider it a change), and if different, we signal a change to the property sheet, which enables the 'Apply' button, for example:

```
            Case WM_COMMAND
                Dim nID As Long = GET_WM_COMMAND_ID(wParam, lParam)
                Dim nCmd As Long = GET_WM_COMMAND_CMD(wParam, lParam)
'...
                ElseIf nID = IDC_EDIT1 Then
                    If nCmd = EN_CHANGE Then
                        hEdit1 = lParam
                        Dim sText As String = GetEditTextW(hEdit1)
                        If sText <> sInitEdit1 Then
                            Debug.Print "Changed Edit1, text=" & sText
                            PropSheet_Changed(GetParent(hDlg), hDlg)
                        End If
                    End If
```
All `PropSheet_` macros and the `GET_WM_COMMAND_` macros defined in Windows headers are provided in tbShellLib, which if you haven't checked it out recently, the API coverage has been expanding quite rapidly; this project for instance requires no additional declarations. The lParam doesn't always contain the hWnd, you can use `GetDlgItem` for those. The demo has a decent selection of controls and different methods for monitoring their changes. 

> [!NOTE]
> You may want to manually track whether something has changed, because if the user clicks 'Ok', or 'Finish' in a wizard, I don't think there's a way to otherwise tell if changes have actually been made.

Finally, for this project, we use the property sheet callback to monitor button presses so we know when to report the changed values:

```
    Private Sub PropsheetCallback(ByVal hwndPropSheet As LongPtr, ByVal uMsg As Long, ByVal lParam As LongPtr)
        Select Case uMsg
'...
            Case PSCB_BUTTONPRESSED
                If lParam = PSBTN_APPLYNOW Then
                    frmMain.PostStatus "Apply Now"
                    ReportValues
                ElseIf lParam = PSBTN_CANCEL Then
                    frmMain.PostStatus "Cancel"
                    frmMain.Text3.Text = "Dialog was canceled; no changes should be made."
```
...and `ReportValues` is just a straight forward function using things like `IsDlgButtonChecked` or cached values from the change messages to provide the final results. 

That's pretty much all there is to it! Check out the demo for the full example code, and download the entire respository for the .res intermediate.

Stay tuned for the next step, turning this into a control panel applet!
 
