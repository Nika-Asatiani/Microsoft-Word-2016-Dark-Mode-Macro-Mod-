# **Word 2016 True Dark Mode Mod 🌙**

This is a simple script (VBA Macro) that forces Microsoft Word 2016 to use a **True Dark Mode** (Black page, White text).

Word 2016 normally only lets you darken the ribbon at the top, leaving the page bright white. This mod fixes that.

## **✨ Features**

* **True Dark Mode:** Turns the page background black and text light gray.  
* **Smart Save Protection:** Automatically turns off Dark Mode for a split second when you save. This ensures your file is saved as a normal white document, so it looks correct when you send it to others.  
* **Print Layout:** Keeps your page margins and layout exactly the same.

## **🛠️ How to Install**

1. Open Microsoft Word 2016\.  
2. Press **Alt** \+ **F11** on your keyboard to open the code editor.  
3. In the list on the left, find **Normal** \-\> **Modules**.  
   * *If you don't see "Modules", right-click "Normal", select Insert \> Module.*  
4. Double-click the Module folder to open a blank white window.  
5. **Copy and Paste** the code below into that window.  
6. Close the code editor window.

## **💻 The Code**

' Helper function to check if we are currently in Dark Mode  
Function IsDarkModeOn() As Boolean  
    Dim doc As Document  
    Set doc \= ActiveDocument  
    ' Check if background visibility is ON and color is BLACK  
    If ActiveWindow.View.DisplayBackgrounds \= True And doc.Background.Fill.ForeColor.RGB \= RGB(0, 0, 0\) Then  
        IsDarkModeOn \= True  
    Else  
        IsDarkModeOn \= False  
    End If  
End Function

' The Main Button Script  
Sub ToggleDarkMode()  
    Dim doc As Document  
    Set doc \= ActiveDocument  
      
    If IsDarkModeOn() Then  
        ' \--- SWITCH TO LIGHT MODE (CLEAN) \---  
          
        ' Turn off background visibility  
        ActiveWindow.View.DisplayBackgrounds \= False  
          
        ' Remove the fill color  
        doc.Background.Fill.Visible \= msoFalse  
          
        ' Reset text to Auto (Black)  
        doc.Content.Font.ColorIndex \= wdAuto  
          
        ' Ensure we are in Print Layout  
        ActiveWindow.View.Type \= wdPrintView  
    Else  
        ' \--- SWITCH TO DARK MODE \---  
          
        ' 1\. Force Word to SHOW the background color  
        ActiveWindow.View.DisplayBackgrounds \= True  
          
        ' 2\. Set background to Black  
        doc.Background.Fill.Visible \= msoTrue  
        doc.Background.Fill.ForeColor.RGB \= RGB(0, 0, 0\)  
        doc.Background.Fill.Solid  
          
        ' 3\. Set text to Light Gray (easier on eyes)  
        doc.Content.Font.Color \= RGB(220, 220, 220\)  
          
        ' 4\. Ensure we stay in Print Layout  
        ActiveWindow.View.Type \= wdPrintView  
    End If  
End Sub

' INTERCEPT SAVE: Runs automatically when you click Save or Ctrl+S  
Sub FileSave()  
    If IsDarkModeOn() Then  
        ToggleDarkMode ' Turn off Dark Mode temporarily  
        ActiveDocument.Save  
        ToggleDarkMode ' Turn Dark Mode back on  
    Else  
        ActiveDocument.Save  
    End If  
End Sub

' INTERCEPT SAVE AS: Runs automatically when you use Save As  
Sub FileSaveAs()  
    If IsDarkModeOn() Then  
        ToggleDarkMode ' Turn off  
        Dialogs(wdDialogFileSaveAs).Show  
        ToggleDarkMode ' Turn back on  
    Else  
        Dialogs(wdDialogFileSaveAs).Show  
    End If  
End Sub

## **🔘 How to Add the Button**

To make this easy to use, add a button to your Word toolbar:

1. Right-click the top menu ribbon in Word and pick **Customize the Ribbon**.  
2. On the right side, click **New Tab** and name it "Mods" (or "Dark Mode").  
3. On the left side, click the dropdown and choose **Macros**.  
4. Click Normal.Module1.ToggleDarkMode and click the **Add \>\>** button.  
5. Click **OK**.

## **🚀 How to Use**

1. Click your new button to turn the lights off. 🌑  
2. Work comfortably.  
3. Press **Ctrl+S** to save. The screen will flicker white briefly—this is normal\! It's cleaning the file so your boss/teacher sees a normal document.