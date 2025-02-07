#Requires AutoHotkey v2.0-a
#Warn  ; Enable warnings to assist with detecting common errors.
;SendMode "Input"  ; Recommended for new scripts due to its superior speed and reliability.

SetTitleMatchMode 1
;F8 to start, F7 to Exit

Pause False


/*###########################################################
Variable initialise##########################################
#############################################################
*/

Global GridWidth := 15 						;Grid width
Global GridHeight := 15 					;Grid Height

Global EndPowerSetting := 20 				;End power setting initialise
Global StartingPowerSetting := 1 			;Starting power setting initialise

Global EndSpeedSetting := 1000 				;End speed setting initialise	
Global StartingSpeedSetting := 100 			;Starting speed setting initialise

Global StartingFrequencySetting := 0.1		; Starting Frequency initialise
Global EndFrequencySetting := 200			; End Frequency initialise

Global MinimumDPMM := 1 ;Minimum value for DPM, dots per mm (Which is the relationship between speed (mm/s) and Frequency (kHz)
Global MaximumDPMM :=5 00 ;maximum Value for DPM, dots per mm
Global DPMM := 15 ;Dots per mm

Global MinimumFrequency := 0.1 ;The maximum & Minimum Settings of the machine, if an out of range or invalid number is given, an error with SAMLight will happen.
Global MaximumFrequency := 500 ; Make sure ALL minimums and maximums are correct, as of this version, this script is unable to tell if
Global MinimumPower := 0.1	;an error has occurred... I doubt that I will put the time in to check for error handling within SAMLight itself
Global MaximumPower := 20
Global MinimumSpeed := 0.01
Global MaximumSpeed := 20000
 
Global SetupType := 0 ;0 for variable power,,,, 1 for constant Power


/*############################################################################################################
Start of Function declarations:::#############################################################################
##############################################################################################################
*/
GridHeightMessage()
{
	Global GridHeight
	msg:=InputBox("Please enter the Vertical count of the test.(Default 15)","Grid Height", "w300 h150", 15)
		if msg.Result="Cancel"
			CANCELLED()
		else
			GridHeight:=msg.Value
			GridWidthMessage()
}

	
GridWidthMessage()
{
	Global GridWidth
	msg:=InputBox("Please enter the Horizontal count of the test.(Default 15)","Grid Width", "w300 h150", 15)
		if msg.Result="Cancel"
			CANCELLED()
		else
			GridWidth:=msg.Value
}

ss(){ 			;Small Sleep duration
	Sleep 50
}
bigs(){ 		;Bigger Sleep duration
	Sleep 75
}
LargeS(){ 		;Large sleep duration used when selecting a new pen
	Sleep 150
}
Reloader(){
	Reload
	Sleep 1000 	;Ensures the App has time to do what it needs before reloading
}

CANCELLED(){
	MsgBox("CANCEL was pressed.")
	Reloader() ;Reloads app if cancel was selected
}
WrongValue(a,b){
	MsgBox("Please enter a value between " %a% " - " %b%)	; takes an input of valid integer ranges and displays them to the user
	Sleep 100												; this will then loop back to the start of the appropriate setting stage
}

StartingMessage(a)
{
	MsgBox("Please ensure that the objects are ordered correctly (X 1-15 -> then upwards 1*Y)")
	Sleep 100
	if a == 1 
	{
	PowerSettingMessage()				; Moves on to a static Power setting mode
	}
	Else StartingPowerSettingMessage()	; Moves on to a static Dots per mm with a variable power setting
}	

PowerSettingMessage(){ ;Runs if SetupType == 1
	Global MinimumPower
	Global MaximumPower
	Global StartingPowerSetting
	msg := InputBox("Please enter the power (W) setting. `n" . MinimumPower . " - " . MaximumPower, "Power Setting", "w300 h150", 20)
		if msg.Result="Cancel"
			CANCELLED()													; Catch for if cancelled has been pressed, then stops the program
			else
			
			If (msg.Value < MinimumPower OR msg.Value > MaximumPower)	; Checks to ensure a valid input is given and that it falls within the specified constraints
			{
				WrongValue(MinimumPower,MaximumPower)					; Gives user that the value given is invalid and provides a valid range
				PowerSettingMessage()
			}
				
			Else
				StartingPowerSetting:=msg.Value							; Sets the starting power to the value given from InputBox
				StartingSpeedSettingMessage()							; Moves on to setting the starting speed setting
				
}

StartingPowerSettingMessage(){ ;Runs if SetupType == 0
	Global MinimumPower
	Global MaximumPower
	Global StartingPowerSetting
	msg:=InputBox("Please enter the power (W) setting to start from. `n" MinimumPower " - " MaximumPower, "Starting Power Setting", "w300 h150", 1)
		if msg.Result="Cancel"
			CANCELLED()													; Catch for if cancelled has been pressed, then stops the program
			else
			
			If (msg.Value < MinimumPower OR msg.Value >= MaximumPower)	; Checks to ensure a valid input is given and that it falls within the specified constraints
			{
				WrongValue(MinimumPower,MaximumPower)					; Gives user that the value given is invalid and provides a valid range
				StartingPowerSettingMessage()
			}
			
			Else
				StartingPowerSetting:=msg.Value							; Sets the Starting power setting to the value given in InputBox
				EndPowerSettingMessage()								; Moves on to setting the End power setting
	}

EndPowerSettingMessage(){
Global EndPowerSetting
Global StartingPowerSetting
Global MaximumPower

	msg:=InputBox("Please enter the end power (W) setting. `n" StartingPowerSetting " - " MaximumPower, "End Power Setting", "w300 h150", 20)
		if msg.Result ="Cancel"
			CANCELLED()															; Catch for if cancelled has been pressed, then stops the program
			else
			
			If (msg.Value <= StartingPowerSetting OR msg.Value > MaximumPower) ; Checks to ensure a valid input is given and that it falls within the specified constraints
			{				
				WrongValue(StartingPowerSetting,MaximumPower)					; Gives user that the value given is invalid and provides a valid range
				EndPowerSettingMessage()
			}
				
			Else
				EndPowerSetting:=msg.Value										; Sets the End power setting to the value provided in InputBox
				StartingSpeedSettingMessage()									; Moves on to setting the Start speed setting
	}

StartingSpeedSettingMessage(){
	Global MinimumSpeed 
	Global MaximumSpeed
	Global StartingSpeedSetting
	
	
	msg:=InputBox("Please enter the speed (mm/s) setting to start from. `n" . MinimumSpeed . " - " . MaximumSpeed, "Starting Speed Setting", "w300 h160", 100)
		if msg.Result="Cancel"
				CANCELLED()													; Catch for if cancelled has been pressed, then stops the program
			else
			
			If (msg.Value < MinimumSpeed OR msg.Value >= MaximumSpeed)		; Checks to ensure a valid input is given and that it falls within the specified constraints
			{																;
				WrongValue(MinimumSpeed,MaximumSpeed)						;
				StartingSpeedSettingMessage()								;
			}
			else
				StartingSpeedSetting:=msg.Value								; Sets the starting speed setting to the value from InputBox
				EndSpeedSettingMessage()									; Moves on to setting the End Speed setting
	}
	
EndSpeedSettingMessage(){
	Global MaximumSpeed
	Global StartingSpeedSetting
	Global EndSpeedSetting
	
	msg:=InputBox("Please enter the end speed (mm/s) setting. `n" StartingSpeedSetting " - " MaximumSpeed,"End Speed Setting", "w300 h150",20000)
		if msg.Result="Cancel"
			CANCELLED()																; Catch for if cancelled has been pressed, then stops the program
		else
			
		If (msg.Value <= StartingSpeedSetting OR msg.Value > MaximumSpeed)			; Check that a valid value is given
		{
			WrongValue(StartingSpeedSetting,MaximumSpeed)
			StartingSpeedSettingMessage()
		}
		Else
			EndSpeedSetting:=msg.Value
		If (SetupType==1)
		{																			; Checks what value is a constant from the Setup value attained from the beginning check
			StartingFrequencySettingMessage()										; Move on to setting the start Frequency
		}
		else
			DPMMMessage()															; Move on to setting the Dots per MM
}

StartingFrequencySettingMessage(){
	Global MinimumFrequency
	Global MaximumFrequency
	Global StartingFrequencySetting
	
	msg:=InputBox("Please enter the Frequency (kHz) setting to start from. `n" MinimumFrequency " - " MaximumFrequency,"Starting Frequency Setting", "w300 h175", 0.1)
		if msg.Result = "Cancel"
			CANCELLED()													; Catch for if cancelled has been pressed, then stops the program
		else
			
		If (msg.Value < MinimumFrequency OR msg.Value >= MaximumFrequency)	; Ensures the starting Frequency is greater than or equal to the minimum limit
		{																; Also checks that it is less than or equal to the maximum limit, and that a valid value is given
			WrongValue(MinimumFrequency,MaximumFrequency)				; If it fails, then WrongValue function is given to address the limits that are valid
			StartingFrequencySettingMessage()
		}	
		else
			StartingFrequencySetting:=msg.Value								; Sets the starting Frequency to the value from InputBox
			EndFrequencySettingMessage()									; Move on to setting the end Frequency value
}

EndFrequencySettingMessage()
{
	Global MaximumFrequency
	Global StartingFrequencySetting	
	Global EndFrequencySetting
	
	
	msg:=InputBox("Please enter the Frequency (kHz) setting to End with. `n" StartingFrequencySetting " - " MaximumFrequency,"End Frequency Setting", "w300 h175", 200)
		if msg.Result="Cancel"
			CANCELLED() 															; Catch for if cancelled has been pressed, then stops the program
		else
			
		If (msg.Value < StartingFrequencySetting OR msg.Value >= MaximumFrequency)	; This block here compares the input value 
		{																		; and ensures the end Frequency is more than the starting Frequency point,
			WrongValue(StartingFrequencySetting,MaximumFrequency)				; but less than or equal to the maximum limit
			EndFrequencySettingMessage()
		}	
		else
			EndFrequencySetting:=msg.Value											; Sets the End point for Frequency to the value from the InputBox
			;GridHeightMessage() ;Incase we want a variable Grid size
			StartMessage()															; Move on to the check to see if SAMLight and the pen setting window is open
}

DPMMMessage(){
	Global DPMM
	msg:=InputBox("Please enter the dots per mm.`nIf unsure use`n33.33 for marking`n333.33 for cutting.`nDots per mm is the relationship between Frequency (kHz) and Speed (mm/s)","Dots Per mm", "w300 h220", 33.33)
		if msg.Result="Cancel"
			CANCELLED()	; Catch for if CANCELLED has been pressed, then stops the program
		else
			DPMM:=msg.Value	; Sets the dots per mm to the value from InputBox
			StartMessage()	; Moves on to checking if pen setting and SAMLight are open
}

	
StartMessage()
{
tText:=""
	If WinExist("SAMLight") ; Checks to see if SAMLight is open (regardless of file that is loaded). If not, there's a catch at the bottom
	{
		MsgBox("Open up the pen #2 settings in SAMLight then press ok on this box to start. LEAVE THE COMPUTER TO DO ITS THING")
		ss()
		Global winID := WinExist("TRUMPF TruPulse nano (SPI-G4) Pen#")	;This block of code ensures that Pen#2 setting exists, then activates the window, carrying on with the task
		ss()															;If not, then the Instruction to the user is repeated
		If WinExist("TRUMPF TruPulse nano (SPI-G4) Pen#")				; Checks to see if Pen Setting window is open, if not, the user is reminded to do so
		{
			WinActivate
			tText:=ControlGetChoice("ComboBox1")						; grabs the current field in the pen number drop down box
			If tText == "2"												; Checks to see if Pen # 2 is selected, if not, it sets it to 2, if it is then proceeds
			{	
				MainFunction()
			}
		
			else 
			{
				ControlChooseString("2", "ComboBox1")	; Proceeds after setting Pen # to 2
				bigs()				
				MainFunction()							; Starts the main loop of assigning values to each pen
			} 
			
		}
		else
		{
			MsgBox("Please ensure you have Pen#2 settings open")	; Reminder to open up the Pen settings window
			ss()
			StartMessage()
		}
	}
	else
	{
		MsgBox("Please ensure SAMLight is running") ; Reminds the user to open up SAMLight and then repeats this block of code
		ss()
		StartMessage()
	}
}



FrequencyFunc(S,D){
Global MinimumFrequency
Global MaximumFrequency
Global ModF										;This function calculates the Frequency when it is tied to Speed during the DPMM constant
	If ((S * D)/1000 >= MaximumFrequency)
	{		;section of this program.
		ModF:= MaximumFrequency					;It also ensures that Frequency can not be outside of the Min & Max
	}
	Else
		If ((S * D)/1000 <= MinimumFrequency)
		{
			ModF:= MinimumFrequency
		}
		else
			ModF:= (S * D)/1000
}

MainFunction(){
	Global StartingSpeedSetting
	Global StartingPowerSetting
	Global StartingFrequencySetting
	Global EndSpeedSetting
	Global EndPowerSetting
	Global EndFrequencySetting
	Global GridWidth
	Global GridHeight
	
	Global ModS := StartingSpeedSetting														;Initialises the main starting variables.
	Global ModP := StartingPowerSetting
	Global ModF := StartingFrequencySetting

	If (EndPowerSetting > StartingPowerSetting)
	{											;Encase of a backwards inputs...
		Global IncrementPower := (EndPowerSetting - StartingPowerSetting) / (GridWidth -1) 	;...This corrects the values of start and end points
	}																					;Given the input constraints, this should never actually be an issue
	Else	
		Global IncrementPower := (StartingPowerSetting - EndPowerSetting) / (GridWidth -1)


	Global IncrementFrequency := (EndFrequencySetting - StartingFrequencySetting) / (GridWidth -1)

	If (EndSpeedSetting > StartingSpeedSetting)	
	{												;Same as above
		Global IncrementSpeed := (EndSpeedSetting - StartingSpeedSetting) / (GridHeight -1)
	}
	Else
		Global IncrementSpeed := (StartingSpeedSetting - EndSpeedSetting) / (GridHeight -1)
	

					

	Loop GridHeight
	{
		ModP := StartingPowerSetting 			;Reset Power setting back to the start
		If (SetupType== 1)
		{										;Checks if we are using a constant Power
			ModF := StartingFrequencySetting	;Reset Frequency back to it's original
		}

		Loop GridWidth
		{
			If (SetupType==0)
			{									;Checks if we are using a constant DPMM
				FrequencyFunc(ModS,DPMM)
			}
		
			ControlSetText ModP, "Edit4", "ahk_class #32770" 	;Power input Field
			ss()												;Small Sleep
			ControlSetText ModS, "Edit5", "ahk_class #32770" 	;Speed Input Field 4
			ss()	
			ControlSetText ModF, "Edit6", "ahk_class #32770" 	;Frequency Input Field 6
			ss()
			ControlClick "Button120", "ahk_class #32770" 		;Apply button highlight 120
			bigs()
			Send "{Enter}"										;Apply button activation
			bigs()
			ControlClick "ComboBox1", "ahk_class #32770" 		;Pen number selection
			bigs()							;Larger Sleep about 3x ss
			Send "{Down}"					;Next Pen
			IsEnabled := ControlGetEnabled("ComboBox1", "ahk_class #32770")	;There's a large delay in SAMLight as it loads the next pen
			while IsEnabled != True
				{
					IsEnabled := ControlGetEnabled("ComboBox1", "ahk_class #32770")
				}
			Send "{Enter}"					;This just removes the pen drop down, probably not needed
			ss()
		
			If (SetupType==0)
			{							;Checks if we're using constant power or constant DPMM
				ModP := ModP + IncrementPower			;Increases Power by increment value
			}
			else
				ModF := ModF + IncrementFrequency 		;Increases Frequency by increment value
		}
		ModS := ModS + IncrementSpeed 					;Increments Speed by Increment value
		ss()
		
	}

}

/*############################################################################################################
End of Function declerations:::######################################################################################
##############################################################################################################
*/
F9::
{
	GridHeightMessage() ; You can change the grid height and width by pressing F9
}

F8::
{ 	
Global SetupType
	Result:= MsgBox("To vary Power(W) and Speed(mm/s) (with a constant DPMM) Press Yes,`nTo vary Speed and Frequency (with a constant Power) Press No`n(Cancel, Exit Program)", "Pen Setup", "YesNoCancel")
	
		If (Result = "Yes")
			{								;This area here will split between two different tests
			SetupType := 0					;Either a constant power with varying frequency and speed (SetupType==1)
			StartingMessage(SetupType)
		}									;or a constant DPMM (ratio between Frequency and speed) with varying speed & power (SetupType==0)
		If (Result = "No")
			{
			SetupType := 1
			StartingMessage(SetupType)
		}
		If (Result = "Cancel")
			{
			ExitApp
		}
	Global ModS := StartingSpeedSetting		;Initialises the main starting variables.
	Global ModP := StartingPowerSetting
	Global ModF := StartingFrequencySetting
}
F7::ExitApp
