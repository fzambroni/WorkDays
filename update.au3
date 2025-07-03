#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.16.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------


#include <InetConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>

; Script Start - Add your code below here
$InetGet = "https://dsext001-eu1-215dsi0708-ifwe.3dexperience.3ds.com/#app:X3DDRIV_AP/content:driveId=DSEXT001&contentId=7F5A7FEDF0DF330068668344001586FE&contentType=file"


        ; Save the downloaded file to the temporary folder.
        Local $sFilePath = @ScriptDir

        ; Download the file in the background with the selected option of 'force a reload from the remote site.'
        Local $hDownload = InetGet($InetGet, $sFilePath, $INET_FORCERELOAD, $INET_DOWNLOADBACKGROUND)

        ; Wait for the download to complete by monitoring when the 2nd index value of InetGetInfo returns True.
        Do
                Sleep(250)
        Until InetGetInfo($hDownload, $INET_DOWNLOADCOMPLETE)

        ; Retrieve the number of total bytes received and the filesize.
        Local $iBytesSize = InetGetInfo($hDownload, $INET_DOWNLOADREAD)
        Local $iFileSize = FileGetSize($sFilePath)

        ; Close the handle returned by InetGet.
        InetClose($hDownload)

        ; Display details about the total number of bytes read and the filesize.
        MsgBox($MB_SYSTEMMODAL, "", "The total download size: " & $iBytesSize & @CRLF & _
                        "The total filesize: " & $iFileSize)


