How to edit this code to suit your own needs.

How to change the MAC Window name;
View form1's code and edit the .MacCaption = "Example1".
Replacing Example1 for the new name.


How to chnage the message that appears when hovering over the tool bar, of the MAC window;
View form1's code and edit the .ToolTipText = "Example2".
Replacing Example2 for the new name

How to chage the Tray Icon;
Go to form 2 and just chnage the forms icon to whatever you want the tray icon to be.


How to chnage the tray menu;
View form2's code and where you see,

mnuRow0.Visible = False
mnuRow1.Visible = False
mnuRow2.Visible = False
mnuRow3.Visible = False
mnuRow4.Visible = False
mnuRow5.Visible = False

Replace the

.visible = False 

to 

.Caption "New Option"

Then go to the bottom of this form and 
input the code below to make the menu active.


Private Sub mnuRow0_Click()

'Put the action that you want to happen when you press your option here.

End Sub



Just keep replacing the number each time you start a new menu option.
Eg: (as there are only 5 MnuRow's that means you can only have 5 options, unless you add more)

This is how it may turn out:


mnuRow0.Caption "Option0"
mnuRow1.Caption "Option1"
mnuRow2.Caption "Option2"
mnuRow3.Caption "Option3"
mnuRow4.Caption "Option4"
mnuRow5.Caption "Option5"


Private Sub mnuRow0_Click()
'Do something
End Sub


Private Sub mnuRow1_Click()
'Do something
End Sub


Private Sub mnuRow2_Click()
'Do something
End Sub


Private Sub mnuRow3_Click()
'Do something
End Sub


Private Sub mnuRow4_Click()
'Do something
End Sub


Private Sub mnuRow5_Click()
'Do something
End Sub