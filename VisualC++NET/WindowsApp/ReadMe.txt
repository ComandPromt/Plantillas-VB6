Windows Forms App with Managed C++
==================================

(c) Richard Grimes 2003

I created this with the New Project wizard. Note that the wizard
creates the first form in the project in files called Form1.h and
Form1.cpp. The first form is called Form1. Furthermore, the wizard
adds the entrypoint code to Form1.cpp which makes little sense. To
tidy up this code here are the steps.

1) remove Form1.h and Form1.cpp from the project through the Solution
Explorer. This will not delete these files.

2) Using Windows Explorer delete Form1.h and Form1.resX

3) Using Windows Explorer rename Form1.cpp to the name of the project
in this case WindowsApp.cpp

4) Using Solution Explorer's Add Existing Item context menu add 
WindowsApp.cpp to the project

5) Now use the Solution Explorer's Add Class context menu to add a
Windows Form (.NET), give this form a meaningful name (I have used
MainForm)

6) When the form is loaded in the IDE it will generate the 
MainForm.resx file

7) Edit the WindowsApp.cpp file to replace all instances of Form1 with
MainForm. There will be two: the name of a header file included at
the top, and in _tWinMain a form instance is created.

