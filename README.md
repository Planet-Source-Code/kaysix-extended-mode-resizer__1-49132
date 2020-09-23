<div align="center">

## Extended mode resizer


</div>

### Description

This is an add-in for adding resizing capabilities to your projects. It is based on the add-in "4 mode Resizer" by PeYTaN (you can find it on

http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=46731&lngWId=1) which is based on the class ControlResizer by Edward Catchpole (you can find it on

http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=46203&lngWId=1)

The (GREAT!!) combined work of the two authors spared me *a lot* of coding time, still I had some complex forms on which i wanted to optimize the resizing to the last twip, and the available modes weren't always enough for me.

So I ended up extending both the class and the add-in to support anchoring and resizing a control with regard to another control, not just to the form.

I also added a small demo project that (I hope!) gives the idea of the various and complex arrangements you can achieve by combining the available modes; just resize the two example forms at will and see how the controls behave.

These are the enhancements:

1. More freedom

Someone commented on the original class that it had the flaw to reserve tag usage for itself, whereas the developer may want to use control tags for their own purposes. This version is tag-friendly: it doesn't replace the previously designed tag, it appends the designed tag after the resizing data and a custom separator.

When the form is loaded and the class initialized, the resizing data are retrieved from the tags and stored in an internal array; the originally designed tags are then restored and no longer needed by the class, as the resizing code will work on the array.

2. More speed

Cycling on the internal array, which only stores info for the controls that will be resized, is faster than cycling everytime on the whole array of form controls, in case at least some controls won't need resizing. Besides, the original class used to set the 4 properties (Left Top Width and Height) separately; this version calculates them

all and touches the control only once, via the Move method. (A SetWindowPos API call might be even faster, in case you want to try it)

3. More options

The new modes allow you to anchor a control to another, so that the gap between those controls stays constant, whereas it grows if both controls grow and move proportionally. A control can now be aligned to another control (Left, Right, Top and Bottom), stretched (Width, Height) to align its right or bottom edge to the latter, or stretched to the same size (Width, Height) of the latter.

Relating a control to another requires that the latter has already been resized, but the modified class takes care of it by sorting the data after assigning a priority to each control, requiring no extra effort from the developer.

Filling and sorting the array will marginally impact on the form startup time, but as said the runtime resizing tends to be faster.

4. More bugs?

Most likely, as there's more code! (I had noticed no bugs in the previous version, just a lacking feature in handling arrays of controls, now properly dealt with) So if you find any, please inform me so that they can be fixed; even better if you send me a fixed version yourself ;)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2003-10-08 15:53:40
**By**             |[KaysiX](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kaysix.md)
**Level**          |Intermediate
**User Rating**    |5.0 (45 globes from 9 users)
**Compatibility**  |VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Extended\_m16563310102003\.zip](https://github.com/Planet-Source-Code/kaysix-extended-mode-resizer__1-49132/archive/master.zip)








