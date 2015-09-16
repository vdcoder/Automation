#include <Date.au3>
#include <Array.au3>

Opt("SendKeyDelay", 50)

; CONSTANTS
$nLogonStartMin = -2 ; Time for logon to start

$sText = InputBox("Carlos Registration", "At what time are postings made available? (Enter hour 1..12)")
If @error <> 1 Then
   $nHour = Int( $sText )
   $sText = InputBox("Carlos Registration", "At what time are postings made available? (Enter minutes 0..59)")
   If @error <> 1 Then
	  $nMin = Int( $sText )
	  $sText = InputBox("Carlos Registration", "At what time are postings made available? (Enter seconds 0..59)")
	  If @error <> 1 Then
		 $nSec = Int( $sText )
		 $sText = InputBox("Carlos Registration", "At what time are postings made available? (Enter am or pm)")
		 If @error <> 1 Then
			If $sText = "am" or $sText = "pm" Then
			   If $sText = "pm" Then
				  If $nHour < 12 Then
					 $nHour = $nHour + 12
				  EndIf
			   Else
				  If $nHour = 12 Then
					 $nHour = 0
				  EndIf
			   EndIf

			   ; Shifts max
			   $sText = InputBox("Carlos Registration", "How many shifts do you want? (Enter 1..5, for all available enter 1000)")
			   If @error <> 1 Then
				  $nShifts = Int($sText)

				  ; Get now time and logon time
				  $sTimeNow = _NowTime( 4 ) ; returns hh:mm
				  $arrParts = StringSplit ( $sTimeNow, ":" )
				  $nNowHour = Int( $arrParts[1] )
                  $nNowMin =  Int( $arrParts[2] )
				  $nLogonHour = $nHour
                  $nLogonMin = $nMin + $nLogonStartMin
				  If $nLogonMin < 0 Then
					 $nLogonMin = $nLogonMin + 60
					 $nLogonHour = $nLogonHour - 1 ; limitation of 12am
				  EndIf
                  If $nLogonHour >= 0 Then

					 ;MsgBox(0, "Wait 1", $sTimeNow & " " & String($nNowHour) & " " & String($nLogonHour) & " " & String($nNowMin) & " " & String($nLogonMin))

					 ; Wait for logon time
					 While $nNowHour < $nLogonHour or $nNowMin < $nLogonMin
						Sleep( 100 )

						; Update now
						$sTimeNow = _NowTime( 4 ) ; returns hh:mm
						$arrParts = StringSplit ( $sTimeNow, ":" )
						$nNowHour = Int( $arrParts[1] )
						$nNowMin =  Int( $arrParts[2] )

						;MsgBox(0, "Wait 1", $sTimeNow & " " & String($nNowHour) & " " & String($nLogonHour) & " " & String($nNowMin) & " " & String($nLogonMin))
					 WEnd

                     ; Automate
                     Login()
					 GoToPage()
					 GetReady()

					 ; Wait for refresh time
					 $sTimeNow = _NowTime( 5 ) ; returns hh:mm:ss
					 $arrParts = StringSplit ( $sTimeNow, ":" )
					 $nNowHour = Int( $arrParts[1] )
					 $nNowMin  = Int( $arrParts[2] )
					 $nNowSec  = Int( $arrParts[3] )
					 While $nNowHour < $nHour or $nNowMin < $nMin or $nNowSec < $nSec
						   Sleep( 100 )

						   ; Update now
						   $sTimeNow = _NowTime( 5 ) ; returns hh:mm:ss
						   $arrParts = StringSplit ( $sTimeNow, ":" )
						   $nNowHour = Int( $arrParts[1] )
						   $nNowMin  = Int( $arrParts[2] )
						   $nNowSec  = Int( $arrParts[3] )
					 WEnd

					 ; Register
                     Register()
				  EndIf
			   Else
				  MsgBox(0, "ERROR", "Midnight internal error")
			   EndIf
			Else
			   MsgBox(0, "ERROR", "Invalid parameter")
			EndIf
		 EndIf
	  EndIf
   EndIf
EndIf

Func Login()
   ;MsgBox(0, "Login", "")

   Send("#r")
   Sleep(2000)
   Send("cmd{ENTER}")
   Sleep(3000)
   Send("""C:\Program Files\Internet Explorer\iexplore.exe"" https://www.yahoo.com{ENTER}")
   Sleep(10000)
   Send("{TAB}{TAB}{TAB}{TAB}{TAB}username{TAB}")

   ; CARLOS PASSWORD ---------------------------------------------------
   Send("password{ENTER}")
   ; CARLOS PASSWORD ---------------------------------------------------

   Sleep(5000)
EndFunc

Func GoToPage()
   ;MsgBox(0, "GoToPage", "")

   ;View Schedule
   Send("{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}")
   Send("{ENTER}")
   Sleep(5000)

   ;Self Schedule
   Send("{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}")
   Send("{ENTER}")
   Sleep(5000)
EndFunc

Func GetReady()
   ;MsgBox(0, "GetReady", "")

   Send("{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}")
EndFunc

Func Register()
   ;MsgBox(0, "Register", "")

   ;Capture
   Send("{SPACE}")
   Sleep(1500)

   ; Monitor source for turn
   $nTrials = 0
   $nPos_Has = 0
   While $nPos_Has = 0 and $nTrials < 10
	  Local $iDelete = FileDelete("c:\temp\f.txt")
	  Sleep(1000)
	  Send("^u")
	  Send("^s")
	  Send("c:\temp\f.txt{ENTER}{ENTER}")
	  Send("^w")
	  $sCapture = FileRead("c:\temp\f.txt")

	  ; Has a turn
	  $sTerm_Has = "</center></td> <td class=""tablewhiteborders""><div align=""left"" class=""smallbodyfont"">"
	  $nPos_Has = StringInStr( $sCapture, $sTerm_Has, 1, 1, 1 )

	  ; Inc trials
	  $nTrials = $nTrials + 1
   WEnd

   If $nTrials = 10 Then
	  MsgBox(0, "Nothing Found", "Nothing Found")
   Else
	  ;Process

	  Local $nTurnIndex = 0
	  Local $aTurnsMonth[1000]
	  Local $aTurnsDay[1000]
	  Local $aTurnsYear[1000]
	  Local $aTurnsStartHour[1000]
	  Local $aTurnsStartMin[1000]
	  Local $aTurnsStopHour[1000]
	  Local $aTurnsStopMin[1000]
	  Local $aTurnsSelect[1000]

	  ; Find for start of first item
	  $nPos = 1
	  $sTerm = "</center></td> <td class=""tablewhiteborders""><div align=""left"" class=""smallbodyfont"">"
	  $nTermSize = StringLen( $sTerm )
	  $nPos = StringInStr( $sCapture, $sTerm, 1, 1, $nPos )
	  While $nPos <> 0
		 ; Advance
		 $nPos = $nPos + $nTermSize

		 ; Read / Parse Date
		 $nTermSize = 10
		 $sDate = StringMid( $sCapture, $nPos, $nTermSize ) ; returns mm/dd/yyyy
		 $arrParts = StringSplit( $sDate, "/" )
		 $nDateMonth = Int( $arrParts[1] )
		 $nDateDay   = Int( $arrParts[2] )
		 $nDateYear  = Int( $arrParts[3] )

		 ; Advance
		 $nPos = $nPos + $nTermSize

		 ; Advance to start time
		 $sTerm = "</div></td><td class=""tablewhiteborders""><div align=""left"" class=""smallbodyfont"">"
		 $nTermSize = StringLen( $sTerm )
		 $nPos = StringInStr( $sCapture, $sTerm, 1, 1, $nPos )

		 ; Advance
		 $nPos = $nPos + $nTermSize

		 ; Read / Parse Start Time
		 $nTermSize = 7
		 $sDate = StringMid( $sCapture, $nPos, $nTermSize ) ; returns hh:mmxm (x is "a" or "p")
		 $arrParts = StringSplit( $sDate, ":" )
		 If StringMid( $arrParts[2], 3, 2 ) = "am" Then
			$nStartHour = Int( $arrParts[1] )
		 Else
			If Int( $arrParts[1] ) < 12 Then
			   $nStartHour = Int( $arrParts[1] ) + 12
			Else
			   $nStartHour = 0
			EndIf
		 EndIf
		 $nStartMin  = Int( StringMid( $arrParts[2], 1, 2 ) )

		 ; Advance
		 $nPos = $nPos + $nTermSize

		 ; Advance to stop time
		 $sTerm = "</div></td><td class=""tablewhiteborders""><div align=""left"" class=""smallbodyfont"">"
		 $nTermSize = StringLen( $sTerm )
		 $nPos = StringInStr( $sCapture, $sTerm, 1, 1, $nPos )

		 ; Advance
		 $nPos = $nPos + $nTermSize

		 ; Read / Parse Stop Time
		 $nTermSize = 7
		 $sDate = StringMid( $sCapture, $nPos, $nTermSize ) ; returns hh:mmxm (x is "a" or "p")
		 $arrParts = StringSplit( $sDate, ":" )
		 If StringMid( $arrParts[2], 3, 2 ) = "am" Then
			$nStopHour = Int( $arrParts[1] )
		 Else
			If Int( $arrParts[1] ) < 12 Then
			   $nStopHour = Int( $arrParts[1] ) + 12
			Else
			   $nStopHour = 0
			EndIf
		 EndIf
		 $nStopMin  = Int( StringMid( $arrParts[2], 1, 2 ) )

		 ; Add to list of turns
		 $aTurnsMonth[ $nTurnIndex ] = $nDateMonth
		 $aTurnsDay[ $nTurnIndex ] = $nDateDay
		 $aTurnsYear[ $nTurnIndex ] = $nDateYear
		 $aTurnsStartHour[ $nTurnIndex ] = $nStartHour
		 $aTurnsStartMin[ $nTurnIndex ] = $nStartMin
		 $aTurnsStopHour[ $nTurnIndex ] = $nStopHour
		 $aTurnsStopMin[ $nTurnIndex ] = $nStopMin
		 $aTurnsSelect[ $nTurnIndex ] = 1
		 $nTurnIndex = $nTurnIndex + 1
	  WEnd

	  ;Delete conflicts
	  $nTurnCount = $nTurnIndex
	  $i = 0
	  While $i + 1 < $nTurnCount
		 If $aTurnsSelect[ $i ] = 1 Then
			$j = $i + 1
			While $j < $nTurnCount
			   If $aTurnsSelect[ $j ] = 1 Then
				  $bRemove = 0
				  ; Same day
				  If $aTurnsMonth[ $i ] = $aTurnsMonth[ $j ] and  $aTurnsDay[ $i ] = $aTurnsDay[ $j ] and $aTurnsYear[ $i ] = $aTurnsYear[ $j ] Then
					 ; i- start in middle
					 If $aTurnsStartHour[ $j ] * 60 + $aTurnsStartMin[ $j ] <= $aTurnsStartHour[ $i ] * 60 + $aTurnsStartMin[ $i ] and $aTurnsStartHour[ $i ] * 60 + $aTurnsStartMin[ $i ] <= $aTurnsStopHour[ $j ] * 60 + $aTurnsStopMin[ $j ] Then
						$bRemove = 1
					 EndIf
					 ; i- stop in middle
					 If $aTurnsStartHour[ $j ] * 60 + $aTurnsStartMin[ $j ] <= $aTurnsStopHour[ $i ] * 60 + $aTurnsStopMin[ $i ] and $aTurnsStopHour[ $i ] * 60 + $aTurnsStopMin[ $i ] <= $aTurnsStopHour[ $j ] * 60 + $aTurnsStopMin[ $j ] Then
						$bRemove = 1
					 EndIf
					 ; j- start in middle
					 If $aTurnsStartHour[ $i ] * 60 + $aTurnsStartMin[ $i ] <= $aTurnsStartHour[ $j ] * 60 + $aTurnsStartMin[ $j ] and $aTurnsStartHour[ $j ] * 60 + $aTurnsStartMin[ $j ] <= $aTurnsStopHour[ $i ] * 60 + $aTurnsStopMin[ $i ] Then
						$bRemove = 1
					 EndIf
					 ; j- stop in middle
					 If $aTurnsStartHour[ $i ] * 60 + $aTurnsStartMin[ $i ] <= $aTurnsStopHour[ $j ] * 60 + $aTurnsStopMin[ $j ] and $aTurnsStopHour[ $j ] * 60 + $aTurnsStopMin[ $j ] <= $aTurnsStopHour[ $i ] * 60 + $aTurnsStopMin[ $i ] Then
						$bRemove = 1
					 EndIf
				  EndIf
				  ; Remove
				  If $bRemove = 1 Then
					 $aTurnsSelect[ $j ] = 0
				  EndIf
			   Endif
			   $j = $j + 1
			WEnd
		 EndIf
		 $i = $i + 1
	  WEnd

	  ;Select
	  Send("{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}")
	  $i = 0
	  While $i < $nTurnCount
		 ; Advance to
		 If $aTurnsSelect[ $j ] = 1 Then
			Send("{SPACE}")
		 EndIf
		 Send("{TAB}")
		 $i = $i + 1
	  WEnd

	  ;Submit
	  If $nTurnCount > 0 Then
		 Send("{ENTER}")
	  EndIf
   EndIf
EndFunc

#cs

#ce
