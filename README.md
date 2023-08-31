# Propsheet Demo

## twinBASIC Property Sheet Demo

This is a demonstration of using dialog resources with the `PropSheetW` API.\
Since twinBASIC has an excellent Form designer for the regular forms tB/VBx users use 99% of the time, there's no built-in dialog designer and no way to convert a form to a dialog resource. So the first step is preparing the resources using another tool. Visual Studio is a good choice, if a little complicated. You basically need to make a new Windows desktop C++ project, use the designer, then compile and recover the .res file from the compiler intermediates outputs. The advantage of this method is you can assign string IDs like IDC_COMMAND1 and it will autogenerate a resource.h file with `#define` statements, which you can copy/paste into tB and do a couple replace operations to make into Consts. Another option, and the one I used for this project, is the excellent tool [ResourceHacker](http://www.angusj.com/resourcehacker/). You can start with an existing .res file, or create a new one.\
For Visual Studio, there's Property Page templates for the dialog editor; start with one of those, from the Resource View tab in the Solution Explorer, where you can right click the .rc and choose 'Add Resource'. It's then got a familiar enough designer you can figure it out from there.\ 
To create a new resource in ResourceHacker, use File->New Blank Script, then Action->Add using script template, then choose the DIALOG template. You can then edit it as text, and/or click the green arrow 'Compile' button to convert the .rc script into a .res, which will bring up the visual editor. Right-click the preview to see options to change properties or insert controls, you can move and size controls in the visual editor. The styles for the dialog you want are `DS_SETFONT | WS_CHILD | WS_VISIBLE | WS_CAPTION | WS_SYSMENU`. Once you've first compiled, you can add image resources from the Action menu, then set the caption to the ID when you insert an ICON or BITMAP control. 

> [!NOTE]
> You can't set the font of controls individually from the resource file, only the default font for the whole dialog and it's controls. The demo project shows you how to set a different font for an individual control at runtime.

After you have your finished resource file, it's time to import it it into twinBASIC. If you've checked out my [UI Ribbon project](https://github.com/fafalone/UIRibbonDemos), you're already familiar with this interim technique until tB can import .res files directly: Use the template .vbp file from this project, edit it in Notepad to change the name of the resource file, then create new tB project via 'Import from VBP'. Note that this will create a Standard EXE by default, but you can change that to a Standard DLL later if needed. 


...
