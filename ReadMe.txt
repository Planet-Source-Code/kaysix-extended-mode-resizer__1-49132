Title: Add-In: Extended mode Resizer
Description: This is an add-in for adding resizing capabilities to your projects. It is based on the add-in "4 mode Resizer" by PeYTaN (you can find it on http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=46731&lngWId=1) which is based on the class ControlResizer by Edward Catchpole (you can find it on http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=46203&lngWId=1)

The (GREAT!!) combined work of the two authors spared me *a lot* of coding time, still I had some complex forms on which i wanted to optimize the resizing to the last twip, and the available modes weren't always enough for me.

So I ended up extending both the class and the add-in to support anchoring and resizing a control with regard to another control, not just to the form.

I also added a small demo project that (I hope!) gives the idea of the various and complex arrangements you can achieve by combining the available modes; just resize the two example forms at will and see how the controls behave.

These are the enhancements:

1. More freedom
Someone commented on the original class that it had the flaw to reserve tag usage for itself, whereas the developer may want to use control tags for their own purposes. This version is tag-friendly: it doesn't replace the previously designed tag, it appends the designed tag after the resizing data and a custom separator.
When the form is loaded and the class initialized, the resizing data are retrieved from the tags and stored in an internal array; the originally designed tags are then restored and no longer needed by the class, as the resizing code will work on the array.

2. More speed
Cycling on the internal array, which only stores info for the controls that will be resized, is faster than cycling everytime on the whole array of form controls, in case at least some controls won't need resizing. Besides, the original class used to set the 4 properties (Left Top Width and Height) separately; this version calculates them all and touches the control only once, via the Move method. (A SetWindowPos API call might be even faster, in case you want to try it)

3. More options
The new modes allow you to anchor a control to another, so that the gap between those controls stays constant, whereas it grows if both controls grow and move proportionally. A control can now be aligned to another control (Left, Right, Top and Bottom), stretched (Width, Height) to align its right or bottom edge to the latter, or stretched to the same size (Width, Height) of the latter.
Relating a control to another requires that the latter has already been resized, but the modified class takes care of it by sorting the data after assigning a priority to each control, requiring no extra effort from the developer. 
Filling and sorting the array will marginally impact on the form startup time, but as said the runtime resizing tends to be faster.

4. More bugs?
Most likely, as there's more code! (I had noticed no bugs in the previous version, just a lacking feature in handling arrays of controls, now properly dealt with)  So if you find any, please inform me so that they can be fixed; even better if you send me a fixed version yourself ;)

...I was almost forgetting...

5. More info
In the previous version, the tag resizing data were a clean, short string of 4 characters representing the resizing modes for Top, Left, Width and Height. I had to resign such a simple syntax as the new Anchor/Align/Stretch modes require more info, like the identity of the control to stick to and eventually the gap in Twips for the anchor mode.

The (totally unreadable, i know!) resizing data are now made like this:
- the 4 parameters are separated by : as their length is no more fixed; what was 0022 is now 0:0:2:2
- a fixed separator is appended to them; I chose a sequence of characters that I won't likely need in the design time tag, or so I hope... if that wasn't your case, change it both in the add-in and in the class
- each parameter can be a number alone or a number followed by _ and a control name; if it's a number below 15, it's an operation mode as it used to be; if 15 or above, it's the amount in twips for the Anchor mode. If the control name is missing, the anchoring is referred to the form.

So, when a control has this tag, 1:4_Label1.0:5_Text1:300§~§TheTag it means that:
Left grows proportionally
Top is aligned to Label1(0).Top
Width grows so that the right border is aligned to the right border of Text1
Height is anchored to the form, it grows always leaving 300 twips from the bottom of the form
The tag is restored to the originally designed value, TheTag

As you will notice from the add-in, modes 4 and 5 have different meanings depending on the parameter you're working on; I know that this adds even more complexity, but as long as the PC running the add-in is not a human (we aren't there yet, are we?!?) I don't really care :)

Further advice: as you see from the example tag above, controls are referred by name (and eventual index) so the tag becomes incorrect if the referred control is renamed/reindexed; in that case the add-in intercepts the error and warns you about it, but it has no way to remap the control under its new name, you will have to choose it manually.
To prevent such inconvenients, my suggestion is to reach a stable configuration of the form and give sensible names to each relevant control before using the add-in to work on resizing.
