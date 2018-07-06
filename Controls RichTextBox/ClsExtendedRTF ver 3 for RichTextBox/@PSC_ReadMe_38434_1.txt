Title: ClsExtendedRTF ver 3 for RichTextBox
Description: ClsExtendedRTF.Cls version 3 for RichTextBox
extensive rewrite, recode and rename.
 
Mostly about manipulating RTF code to change text and background colours.
ClsExtendedRTF itself, is now about file and doc level activities
ClsRTFFontPainter contains all the RTF font manipulation stuff.
Now includes API and RTF based higlighting(Not selection, this is the RichTextBox 
equivalant of highlihter pens). 
Highlighting with:
API advantage it can detect highlighting; disadvantage single colour at a time
RTF advantage multicolour highlighting. disadvantage can't detect itself.
ClsAPIZoom for RichTextBox, a few lines of code and your RTBox is zoomable.
cLsManifestation (incorperated from my other upload) Gives compiled program user's choice
of Classic or WindowsXP(if they have XP)
Added panels(form) which give you greater control over highlight, text colour and text format.
*New* Materials interface you can create your own materials colour schemes.
*New* Styles; if you create a text and background colour scheme you want to reuse
 save it with a name and access through the RTF Font Painter tool and class.
fixed small bug in RemoveFormatting which added a space to start;(if selection was at start of doc)
This file came from Planet-Source-Code.com...the home millions of lines of source code
You can view comments on this code/and or vote on it at: http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=38434&lngWId=1

The author may have retained certain copyrights to this code...please observe their request and the law by reviewing all copyright conditions at the above URL.
