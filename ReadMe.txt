I wrote this little AddIn because I have blocks of code that I use fairly often such as the error handling code. To save me having to type this all the time I wrote this AddIn. I think that this whole AddIn is about as simple as an addin can get and as such is maybe the best learning example for a VB IDE AddIn. I have also added some comments to help you understand how I determined certain properties for my own items.
If you want to create new menu items you will have to create a new MenuHandler variable to recieve the menuitem_click event, edit the code by copying the three lines of code to create your own menu item and then create the copy the code for the menu handler event. The following blocks are the ones that would be copied:

		Public WithEvents YourMenuHandler As CommandBarEvents

		Set cbCmdBarCtrl = mcbMenuCommandBarCtrl.Controls.Add(1)
		cbCmdBarCtrl.Caption = <Caption on Menu>
		Set Me.YourMenuHandler = VBInstance.Events.CommandBarEvents(cbCmdBarCtrl)

		Private Sub YourMenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
		'Event runs when menu item is clicked
		On Error Resume Next
		   Dim strText As String
		   '
		      strText = GetStringFromFile(<nameof section>)
		      Call VBInstance.ActiveCodePane.CodeModule.InsertLines(<LastLine|FirstLine|CurrentStartLine>, strText)
		End Sub

I have tried to make this application configurable through text files but as yet have not gotten it to work. The reason for this is that I have not yet figured out how to assign multiple menu items to one MenuHandler. If I can work this out I will release a second version. So if anyone knows the answer please let me know. If we can get that to work then it will be possible to place the menu items in a text file and never have to edit the AddIn to add or remove menu items.

Please Check out my other published source code on www.planet-source-code.com or visit my website www.dazzlingsoftware.com where you will find everything I have published so far.
Thanks
Darryl Hasieber
Dazzling Software
Contact Details:
WebSite: 	http://www.dazzlingsoftware.com
E-Mail:		darrylha@dazzlingsoftware.com