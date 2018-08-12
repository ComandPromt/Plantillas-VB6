
 How to use the RES Protocol of the Webbrowser Control in VB
 ========================================================================
 by Andreas Schwarz , andi@futureprojects.de

 
 Here are 5 steps for creating a RES Interface for the Webbrowser Control in VB:

 1. Create a Resource File , using a RC-Script and the RC.EXE Resource Compiler
    (view in Resource Example Directory)

 2. Include the created resourcefile in your project
 3. Include the Webbrowser Control in your project
 4. When you will display a page use this command


    Private Sub DisplayResAdress(ResourceName as String,wb as Webbrowser)
         wb.navigate "res://" & app.path & "\" & app.exename & ".exe/" & ResourceName
    End Sub

 5. Show the results...


 Ready & go ...

 bye
 Andi
 