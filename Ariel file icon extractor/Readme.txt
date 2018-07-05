----------------------------------------------------------------
Ariel Associated File Icon Extractor
Description: Test program to demo icon extraction
Release    : VB6 SP4
Copyright  : © T De Lange, 2000
E-mail     : tomdl@attglobal.net
----------------------------------------------------------------
This project demonstrates how to extract icons associated with
files into an imagelist and displaying them in a listview with
the filenames.
The SHGetFileInfo function of the shell32.dll library is used,
which makes the job much easier than before. The ImageList_Draw
function in comctl32.dll is used to draw the icon in a picture box,
from where it is placed into the image list.

Watch out for the following:
a) Image list can hold only approx 400 icons, so you will have
   to remove duplicate images for files other than exe's
b) Remember to set the lvw's mask color to the appropriate
   system color, usually buttonface.
----------------------------------------------------------------
Credits:
Peter Meier, Planet Source Code for
the technique as used in his 'DelRecent' posting
----------------------------------------------------------------
