
This form contains an example of how to do a simple Outlook-type split form.
The splitter code is modifed from the http://www.vb-helper.com/HowTo/hsplit.zip code.

The buttons on the left pane are in a control array so to add more buttons, copy
one of the existing buttons and paste into the left-pane picture box. That is,
make sure you select the left-pane picture box first before pasting.

Similarly, the icons are also displayed in a picture control array. I am using a resource
file to store the icons and load them dynamically into the picture control array when
a button is clicked. Each icon in the resource file in indexed by a string-identifier and
in the form, a 2-dimension string array keeps track of the active icons for each button.
If an icon is inactive for a button, then set the corresponding element in the string
array to empty. The Form_Load event contains a section which initialises the string
array with icon index values for each button in the form. Therefore, to add more icons,
you would simply need to add them to the resource file and rename the index to a suitable
string name. Edit the Form_Load event and assign appropriate element with the new icon
index string.

Labels for each icon are also displayed using a label array control.

If icons are obscure by resizing, the appropriate scroll button should appear allowing
users to scroll up or down to see the obscured icons.

The code for Make3D is borrowed and customized from the Make3d code submitted by Matthew
Inman. (Thanks Matthew)

Email me at hqvu@totalise.co.uk if you require further info.

Have fun.

