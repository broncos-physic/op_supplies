#Include <GeneralFunctions>
#Include <GUIs>
#Include <PharmNOVAFunctions>
#Include <PharmNetFunctions>
#Include <PowerChartFunctions>
#Include <SQLFunctions>
#Include <Calculators>
#MaxThreads 20

; ListLines False
#Requires AutoHotkey v2.0
#SingleInstance force
Critical "Off"                  ; To ensure script is interruptible.
CoordMode "Mouse", "Window"     ; Define mouse positioning relative to window.
CoordMode "Pixel", "Window"     ; Define pixel (including ImageSearch) area relative to window.
SetTitleMatchMode "RegEx"       ; RegEx quick reference: https://www.autohotkey.com/docs/v2/misc/RegEx-QuickRef.htm

/*
Autohotkey v2 must be installed on user's computer.  TRM approved:  https://trm.oit.va.gov/ToolPage.aspx?tid=6458#  Download available:  https://www.autohotkey.com/
Stores scripted hotkeys to be used for automating processes in Oracle Cerner pharmacy applications.

Contact: VHAPharmNOVA@va.gov
Additional documentation: https://dvagov.sharepoint.com/sites/PharmacyCernerEHRCommunityofPractice/SitePages/Oracle-Cerner-Repository-for-Automating-Pharmacy-VA.aspx
Project repo: https://github.ec.va.gov/EPMO/Oracle-Cerner-Repository-for-Automating-Pharmacy-VA
*/

/************************************************************ Auto-Execute Thread *****************************************************************
; This is the stuff that will run when the script is turned on.
***************************************************************************************************************************************************/

; Define OnError (logs error to file)
OnError LogError

; Checks for icon presence in script directory, changes icon to it if exists.
If (FileExist(A_ScriptDir "\Images\PharmNOVA.ico"))
    {
    IconPath := A_ScriptDir "\Images\PharmNOVA.ico"
    TraySetIcon IconPath,,true
    }

Try
    {
    PAFHandle := LoadPicture(A_ScriptDir "\Images\PAF.bmp")
    PANHandle := LoadPicture(A_ScriptDir "\Images\PAN.bmp")
    PAVHandle := LoadPicture(A_ScriptDir "\Images\PAV.bmp")
    NFHandle := LoadPicture(A_ScriptDir "\Images\NF.bmp")
    FormRestHandle := LoadPicture(A_ScriptDir "\Images\FR.bmp")
    FormHandle := LoadPicture(A_ScriptDir "\Images\F.bmp")
    }

Try
    {
    ; Test query to check LSV permission - array values will only exist if query is successful
    Admin := StrSplit(LSVQuery(Format("SELECT UserAdmin, DrugFileAdmin, HotstringAdmin, HotkeyAdmin, Email, BetaTester FROM LSV.D05_VISN20INFORM.PharmNOVA_Permissions WHERE SamAccountName = '{1}'",A_UserName),0,5,3),'`t')
    global UserAdmin := RegExReplace(Admin[1],'\D')
    global DrugFileAdmin := RegExReplace(Admin[2],'\D')
    global HotstringAdmin := RegExReplace(Admin[3],'\D')
    global HotkeyAdmin := RegExReplace(Admin[4],'\D')
    global UserEmail := Trim(Admin[5])
    global BetaTester := RegExReplace(Admin[6],'\D')
    global LSVAccess := 1
    }
Catch
    {
    global UserAdmin := 0, DrugFileAdmin := 0, HotstringAdmin := 0, HotkeyAdmin := 0, LSVAccess := 0, UserEmail := "", BetaTester := 0
    }

; Define global variables.
global Version := A_Now, xPosRxItem := "", yPosRxItem := "", xPosEligGUI := "", yPosEligGUI := ""

; Define global variables for images used in ImageSearch.
; File path + options for images used to check window for drop down location or application status.
; see https://www.autohotkey.com/docs/v2/lib/ImageSearch.htm
ImageMagnifyingGlass := LoadPicture(A_ScriptDir "\Images\MagnifyingGlass.bmp")             ; Magnifying glass used to open search window (prescriber in MMR specifically).
ImageReady := LoadPicture(A_ScriptDir "\Images\Ready.bmp")                                 ; MMR/WQM, lower left corner = "Ready" (records exist, app is idle)
ImageLoading := LoadPicture(A_ScriptDir "\Images\WQMLoading.bmp")                          ; Work Queue Monitor, lower left corner = "Loading" (batch action is completing)
ImageNoRecords := LoadPicture(A_ScriptDir "\Images\WQMNoRecords.bmp")                      ; Work Queue Monitor, lower left corner = "No records found" (no records, app is idle)
ImageRoutingOptionReason := LoadPicture(A_ScriptDir "\Images\WQMRoutingOptionReason.bmp")  ; Work Queue Monitor, column header.
ImageDoNotMail := LoadPicture(A_ScriptDir "\Images\RoutingOptionValidationDoNotMail.bmp")  ; Routing Option Validation Reason window, "Do Not Mail" status in Patient Mailing Preference section
ImageOrderSentence := LoadPicture(A_ScriptDir "\Images\MMROrderSentenceHeader.bmp")        ; Order Sentence column header in Amb Profile of MMR.
ImageRxNumHeader := LoadPicture(A_ScriptDir "\Images\RxNumHeader.bmp")                     ; Rx Number header in WQM or MMR.
ImagePCCustomInfoMRN := LoadPicture(A_ScriptDir "\Images\PCCustomInfoMRN.bmp")             ; MRN field in Custom Information Window.
ImagePCQuickSearch := LoadPicture(A_ScriptDir "\Images\PCQuickSearch.bmp")                 ; Magnifying glass + dropdown for patient search in PowerChart.
;ImageRxRoutingComplete := LoadPicture(A_ScriptDir "\Images\RxRoutingOptionComplete.bmp")   ; Used to confirm drop down has closed and script can proceed.
ImageRxDateTime := LoadPicture(A_ScriptDir "\Images\RxDateTime.bmp")                       ; Dispense Priority Date/Time
ImageWQMHorizontalScroll := LoadPicture(A_ScriptDir "\Images\WQMHorizontalScroll.bmp")     ; Horizontal scroll bar in Work Queue Monitor.
ImageNom := LoadPicture(A_ScriptDir "\Images\Nom.bmp")                                     ; Nominate providers for access to prescribe controlled substances.
ImageBegin := LoadPicture(A_ScriptDir "\Images\Begin.bmp")                                 ; Begin nomination process.
ImageNoResultsFound := LoadPicture(A_ScriptDir "\Images\NoResultsFound.bmp")               ; No providers to nominate for station.
ImageNomProviders := LoadPicture(A_ScriptDir "\Images\NomProviders.bmp")                   ; Providers to nominate.
ImageDEA := LoadPicture(A_ScriptDir "\Images\DEA.bmp")                                     ; Rx DEA field.
ImageBirthday := LoadPicture(A_ScriptDir "\Images\birthday.bmp")                           ; MMR birthday in ambulatory profile... double-click to copy!
ImageNoMatch := LoadPicture(A_ScriptDir "\Images\NoMatch.bmp")                             ; No providers matched to display.
ImageNoNPI := LoadPicture(A_ScriptDir "\Images\MissingNPI.bmp")                            ; Provider missing NPI on file.
ImageAddAlias := LoadPicture(A_ScriptDir "\Images\AddAlias.bmp")                           ; New prsnl alias icon.

; AHK documentation suggests using groups to optimize #HotIf, so creating groups to define context for hotkey.
; Prefer using RegEx terms instead of groups, removed some groups 08/16/2024
; Alt+left-click uses this group
GroupAdd "MSFT", "Microsoft Teams"
GroupAdd "MSFT", "- Excel"
GroupAdd "MSFT", "- Outlook"                                                                ; Outlook via reading pane
GroupAdd "MSFT", "- Message"                                                                ; Open outlook message
GroupAdd "MSFT", "ahk_class Chrome_WidgetWin_1",,"Azure Data Studio|Visual Studio Code"     ; Edge or Chrome browser, exclude ADS and VS Code

; Alt+Shift+P (Pharmacy Eligibility) uses this group
GroupAdd "Elig", "Pharmacy Eligibility"                                                     ; Results window
GroupAdd "Elig", "Person Search - \\Remote"
GroupAdd "Elig", "PharmNet: Ret"

; Set timer to reload script between 00:00-00:59 if user has been inactive for at least 60 minutes.
SetTimer ScheduledMaintenance, -1000

; Capture monitor characteristics
; Some scripts attempt to place or move a window based on another window, so we need to know the display boundaries.
MonitorWidth()

; Check that user has current launcher version saved to workstation.
CheckLauncher()

; Check for existing ini file and create if not found.
; If ini exists, use it to set variables.
; This sets Sta3n, so must run before LDFUpdate.
CheckIni(0)

If !RegExMatch(Sta3n,"^\d+$")
    Sta3n := 0

If BetaTester AND Sta3n AND Sta3n > 999
    {
    global UserAdmin := 1, DrugFileAdmin := 1, HotstringAdmin := 1, HotkeyAdmin := 1
    }
Else
    global BetaTester := 0

; Modify options available from system tray
Try if !HotkeyAdmin
    {
    A_TrayMenu.Insert("&Help","&User Menu", (*) => GUIMain())
    A_TrayMenu.Insert("&User Menu","&Help (SharePoint)", (*) => Run("https://dvagov.sharepoint.com/:u:/r/sites/PharmacyCernerEHRCommunityofPractice/SitePages/Oracle-Cerner-Repository-for-Automating-Pharmacy-VA.aspx?csf=1&web=1&share=EZYFbCtdCr9Bvseh7gKT8tMB4MMSjq_qW1N5PuvnfEAATw&e=o2SoTo"))
    ; A_AllowMainWindow := False ; Disabling this makes Alt+Shift+X error reporting fail so remove Open menu instead
    A_TrayMenu.Delete("&Open")
    A_TrayMenu.Delete("&Help")
    A_TrayMenu.Delete("&Edit Script")
    A_TrayMenu.Delete("&Window Spy")
    A_TrayMenu.Default := "&User Menu"
    }

Try if HotkeyAdmin
    {
    A_TrayMenu.Insert("&Help","&User Menu", (*) => GUIMain())
    A_TrayMenu.Default := "&User Menu"
    A_TrayMenu.Insert()
    A_TrayMenu.Insert(,"Review &Error Log", (*) => Run(A_ScriptDir "\OCRAP Error Log.xlsx"))
    A_TrayMenu.Insert(,"&View Station Folders", (*) => Run(A_ScriptDir "\Station"))
    }

; Check for Local Drug File reference, and load to map if exists.
; This should run after CheckIni (populates Sta3n).
; This defines LDFFile and LDF items so must run before GUI.
LDFUpdate()

; Hotstring globals.
HotstringRows := []
HotstringList := []
HotstringManagerOpen := false
HotstringEditorOpen := false
HotstringRepopen := false
HotstringShopen := false

OpenPREVENT := false ;PREVENT global.
OpenCrCl := false ;CrCl global.
SearchGUIopen := false ;SearchGUI global.
PatientLookupGUI := false ;Must close GUIEligibility() before recalling Eligibility()... unsure why, but started double-triggering in Sep. '26.

; Initialize hotstring CSV file, or disable functionality and hide editor if it cannot be found/created/opened.
; Must get IniDir from CheckIni first.
HotstringFile := EnsureHotstringcsvPath()
If HotstringFile
    InitHotstringCSV()

Busy := false

; LOCAL HOTKEYS
If !BetaTester AND FileExist(A_ScriptDir "\Site Tools\" Sta3n ".ahk")
    Run A_ScriptDir "\Site Tools\" Sta3n ".ahk"

; Display GUI!
GUIMain()

; ListLines controls what is recorded in log.
ListLines True

/************************************************************ Alt+O (OK) ***************************************************************************
; This hotkey will check for presence of a DoD HP and attempt to remove or replace using CheckHP() function above.
; NOTE: Scripts that include "!o" at these windows may be adversely impacted. Additional controls and/or modification of (#)HotIf logic may be necessary.
;   https://www.autohotkey.com/docs/v2/lib/_HotIf.htm
;   Expected outcome would be that initial hotkey would interrupt, this hotkey will run, then prior hotkey will resume.
;   This may be acceptable in this instance and would just need to ensure action is not duplicated.
;   Initial testing of failures hotkey provides no evidence of duplication
;   Consider using A_Priorhotkey: https://www.autohotkey.com/docs/v2/Variables.htm#PriorHotkey
; NOTE: "!o" will not trigger itself "By default, a given hotkey or hotstring subroutine cannot be run a second time if it is already running."
;   https://www.autohotkey.com/docs/v2/misc/Threads.htm
****************************************************************************************************************************************************/

; #HotIf does not prevent other hotkeys from calling this hotkey
; Hotkey will not call itself
#HotIf WinActive("i)^(Refill|New Order|Bill Only).* for Patient:.* - \Q\\Remote\E")
!o::
{
KeyWait "Alt", "T2"
If CheckHPOn = 1
    {
    RxCheckHP()
    Send "!o"
    }
}

/************************************************************ Ctrl+D (DEA) ************************************************************************
; Check formatting and validity of DEA numbers in the Rx finishing workflow.
; Members of "VHAPBM Pharmacists (ALL)" AD group with LSV access will be able to query against national registry, too, secured per specifications from T1SMET.
****************************************************************************************************************************************************/

^d::
{
KeyWait "Control", "T2"
If WinActive("^New Order for Patient.* - \Q\\Remote\E")
    {
    If ColorFilter = 2
        {
        Send "#^c" ;Disable color filter.
        Sleep 500
        }
    Loop 3
        {
        Sleep 250
        } until ImageSearch(&X, &Y, 7, 84, 722, 455, "*75 HBITMAP:*" ImageDEA) ;Find DEA field.
    If ColorFilter = 2
        Send "#^c" ;Enable color filter.
    If X = "" OR Y = ""
        {
        Tooltip "DEA not found. One may not be required.", 881, 32
        SetTimer () => ToolTip(), -3000
        Exit
        }
    MouseClick(,X+102,Y+26)
    ;Copy
    A_Clipboard := ""
    Loop 3
        {
        Send "^c"
        If Clipwait(1)
            Break
        }
    If A_Clipboard = "" OR A_Clipboard = "(None)"
        {
        InputDEA := InputBox("DEA field appears empty. Please enter the DEA number.`n`nCancel to exit.","Enter DEA","h150")
        If InputDEA.Result != "OK" OR !RegExMatch(InputDEA.Value, "\S+")
            Exit
        DEA := RegExReplace(InputDEA.Value, "\s*")
        }
    Else
        DEA := RegExReplace(A_Clipboard, "\s*")
    }
Else
    {
    InputDEA := InputBox("DEA field appears empty. Please enter the DEA number.`n`nCancel to exit.","Enter DEA","h150")
    If InputDEA.Result != "OK" OR !RegExMatch(InputDEA.Value, "\S+")
        Exit
    DEA := RegExReplace(InputDEA.Value, "\s*")
    }
CheckDEA(DEA)
}
#HotIf

#HotIf WinActive("^(Search Physician|Define Aliases).*Remote$")
^d::
{
KeyWait "Control", "T2"
InputDEA := InputBox("Please enter the DEA number.`n`nCancel to exit.","Enter DEA","h150")
If InputDEA.Result != "OK" OR !RegExMatch(InputDEA.Value, "\S+")
    Exit
DEA := RegExReplace(InputDEA.Value, "\s*")
CheckDEA(DEA)
}
#HotIf

#HotIf Sta3n = 668 AND WinActive("^(Search Physician|Define Aliases|New Order).*Remote$")
^e::
{
KeyWait "Control", "T2"
InputNPI := InputBox("Please enter the NPI number.`n`nCancel to exit.","Enter NPI","h150")
If InputNPI.Result != "OK" OR !RegExMatch(InputNPI.Value, "(1\d\d\d\d\d\d\d\d\d)", &NPI)
    {
    If InputNPI.Result = "OK"
        MsgBox("NPI not found or invalid. Please try again.","Invalid/Missing NPI","4096 T4")
    Exit
    }
CheckLEIE(NPI[1])
}
#HotIf

#HotIf WinActive("^E-Prescription.*Remote$")
^d::
{
KeyWait "Control", "T2"
If !WinActive("^E-Prescription Detail.*Remote$")
    {
    If !WinActive("Authentic EPCS – received via DEA") AND A_Now > 20251107120000 ;Scheduled go-live for defect patch.
        {
        EPCS := MsgBox("It appears this eRx may not be authentic, or received via DEA-approved processes.`n`nProceed anyway?","eRx Authenticity Check","YN 4096")
        If EPCS != "Yes"
            Exit
        WinMatch("E-Prescription",3)
        }
    Send "!d" ;Order Details
    WinMatch("E-Prescription Detail - \\\\Remote",3)
    }
Try
    eRx := ERxDetails()
Catch
    {
    ToolTip "Failed to activate E-Rx Prescription Details window. Please try again."
    SetTimer () => ToolTip(), -3000
    }
Try
    {
    RegExMatch(eRx["Prescriber DEA"],"([A-Za-z]{2}\d{7})",&DEA)
    CheckDEA(DEA[1])
    }
Catch
    {
    InputDEA := InputBox("DEA not found. Please enter the DEA number.`n`nCancel to exit.","Enter DEA","h150")
    If InputDEA.Result != "OK" OR !RegExMatch(InputDEA.Value, "\S+")
        Exit
    DEA := RegExReplace(InputDEA.Value, "\s*")
    CheckDEA(DEA)
    }
}
#HotIf

#HotIf Sta3n = 668 AND WinActive("^E-Prescription.*Remote$")
^e::
{
KeyWait "Control", "T2"
Send "!d" ;Order Details
WinMatch("E-Prescription Detail - \\\\Remote",3)
Try
    eRx := ERxDetails()
Catch
    {
    ToolTip "Failed to activate E-Rx Prescription Details window. Please try again."
    SetTimer () => ToolTip(), -3000
    }
Try
    {
    Try
        FirstName := Trim(eRx["Prescriber First Name"])
    Catch
        FirstName := ""
    Try
        LastName := Trim(eRx["Prescriber Last Name"])
    Catch
        LastName := ""
    Try
        ZipCode := Trim(eRx["Prescriber Zip Code"])
    Catch
        ZipCode := ""
    Try
        State := Trim(eRx["Prescriber State"])
    Catch
        State := ""
    Try
        City := Trim(eRx["Prescriber City"])
    Catch
        City := ""
    Try
        Street := Trim(eRx["Prescriber Street"])
    Catch
        Street := ""
    RegExMatch(eRx["Prescriber NPI"], "(1\d\d\d\d\d\d\d\d\d)", &NPI)
    CheckLEIE(NPI[1], LastName, FirstName, ZipCode, State, City, Street, false)
    }
Catch
    {
    InputNPI := InputBox("Please enter the NPI number.`n`nCancel to exit.","Enter NPI","h150")
    If InputNPI.Result != "OK" OR !RegExMatch(InputNPI.Value, "(1\d\d\d\d\d\d\d\d\d)", &NPI)
        {
        If InputNPI.Result = "OK"
            MsgBox("NPI not found or invalid. Please try again.","Invalid/Missing NPI","4096 T4")
        Exit
        }
    CheckLEIE(NPI[1])
    }
}
#HotIf

/********************************************** Rx Item Attributes ****************************************************************
; These hotkeys serve to provide a summary of Rx item (drug file) attributes during prescription processing
; This leverages multiple functions, including LSVQuery so LSV access is necessary for this to work (expect try/catch and other errors
; should allow script to exit gracefully if failure encountered)
; WinMoveRxAttributes is used to move the MsgBox, default location is center of screen, which is unfortunate. Function allows a new thread
; to open, complete task, then resume prior thread (i.e. MsgBox timeout)
; Product Selection window can be exited 4 ways with item selection: Enter, left-click OK, double-click item, Alt+O (OK)
;   Attempting to "intercept" those 4 actions and grab item attributes prior to window closing
;   If Mail Order (CMOP) is checked (= true), then window is closed and Rx item query is run, if no results, no window is expected.
;   Alt+O and Enter should be okay; left-click is very problematic so initially using right-click as a substitute (right-click on OK
;   or double-right-click)
;
; When updated to GUI/List View, will try to leverage these data sources as well:
; SELECT distinct producttext FROM LSV.[PBM_FieldCollab].[PPSN_AllProducts] WHERE PRODUCTTEXT IS NOT NULL ORDER BY PRODUCTTEXT
; select top 10 * from LSV.[PBM_FieldCollab].[vPBM_NationalRpts_VANFapp_CFU]
; select top 10 * from LSV.[PBM_FieldCollab].[vPBM_NationalRpts_VANFapp_UEformulary]
**********************************************************************************************************************************/
/*
; Enter and Alt+O can be used a lot of places so #HotIf
; \.*?+[{|()^$ must be escaped for RegEx, "\QLiteral Text\E" escapes full string
#HotIf WinActive("\QProduct Selection - \\Remote\E") AND RxItemPosition ~= "[1-3]"
Enter::
!o::
{
; Check if user has already loaded NDCArray via LButton hotkey
; If yes, send keys and let other thread handle it
; Possible gap if person clicks row, navigates with up/down arrow, hits enter
; If NDCArray.Has(-1)
If NewRxNDC
    {
    Send (A_ThisHotkey = "!o" ? A_ThisHotkey : "{Enter}")
    Return
    }

; If NewRx capture clipboard for restore and update flag
If NewRx
    global NewRx := false, NewRxNDC := "", RxItemPriorClip := ClipboardAll()

A_Clipboard := ""
Send "^c"
ClipWait(0.25)

If FoundPos := RegExMatch(A_Clipboard,"\t\d{5}-\d{4}-\d{2}\t",&Match)
    global NewRxNDC := SubStr(A_Clipboard,Match.Pos+1,Match.Len-2)
Else A_Clipboard := RxItemPriorClip

; Send trigger after value extracted
Send (A_ThisHotkey = "!o" ? A_ThisHotkey : "{Enter}")

; Enter might be used to perform search and not just selection
; So, if Prod Selection window doesn't close or we didn't capture an NDC, exit rather than query
If !WinWaitClose("\QProduct Selection - \\Remote\E",,1) OR !IsSet(ID)
    {
    A_Clipboard := RxItemPriorClip
    Return
    }

; If result window already exists, close it then run query
If WinExist("Item Details for")
    {
    WinClose("Item Details for")
    }

RxItem(ID,"NDC")
}

; 'Passive' Left-Button NDC capture
; Disabled 03/20/2025 because I haven't figured out a way to consistently capture the NDC value AND avoid replacing clipboard contents
; Clipboard restore is much more reliable with this option, but double-click fails to capture NDC with this option
#HotIf WinActive("\QProduct Selection - \\Remote\E")
~LButton::{
; We put this here instead of HotIf because we don't want it to block other hotkeys that use mouse at this window
If NOT RxItemPosition ~= "[1-3]"
    Return
If NewRx
    global NewRx := false, NewRxNDC := "", RxItemPriorClip := ClipboardAll()
;    global NewRx := false, NDCArray := [], RxItemPriorClip := ClipboardAll()

KeyWait "LButton", "U"

; Clipwait only controls script flow, it doesn't cancel copy action, etc.
; When a user clicks directly on OK, the copy command is sent but the Product Selection window is closed immediately
; This can result in the clipboard being restored prematurely, then replaced with Rx Qty (or other undesired value)
; Tried using WinGetPos and MouseGetPos but this seems to introduce too much lag preventing reliable clipboard capture
; This scenario is what we would expect if user clicks OK/cancel/X or on second click to select (first click should be captured)
A_Clipboard := ""
Send "^c"
While !ClipWait(0.25)
    {
    If !WinExist("\QProduct Selection - \\Remote\E")
        {
        WinWaitActive("New Order for Patient",,10)
        If NOT NewRxNDC
            {
            A_Clipboard := RxItemPriorClip
            global NewRx := true, NewRxNDC := ""
            }
        Return
        }
    }

If FoundPos := RegExMatch(A_Clipboard,"\t\d{5}-\d{4}-\d{2}\t",&Match)
    global NewRxNDC := SubStr(A_Clipboard,Match.Pos+1,Match.Len-2)

; Timer period should be shorter than ClipWait period
SetTimer ProductSelected, 200

ProductSelected(){
; If product selection is closed and we don't have an NDC
If !WinExist("Product Selection") AND NOT NewRxNDC
    {
    global NewRx := true
    If WinWait("i)^E-Prescription",,5)
        WinWait("New Order for Patient",,10)
    A_Clipboard := RxItemPriorClip
    SetTimer ProductSelected, 0
    Return
    }

; If product selection window has closed, and we have an NDC, then try query
; Global variables are reset via RxItem when either
;   1) New Order window fails to load within 30 seconds
;   2) Item Details window fails to load within 30 seconds
;   3) Item Details window successfully loads
Else if !WinExist("Product Selection")
    {
    Try if NewRxNDC
        RxItem(NewRxNDC,"NDC")
    Else
        {
        A_Clipboard := RxItemPriorClip
        global NewRxNDC := "" ; , RxItemPriorClip := ""
        }
    Return
    }
}
}
#HotIf

#HotIf


/* Initial version
~LButton::{
If !IsSet(RxItemPosition) OR RxItemPosition = 4 OR !WinExist("\QProduct Selection - \\Remote\E")
    Return
If NewRx
    global RxItemPriorClip := ClipboardAll(), NewRx := false, NDCArray := []

KeyWait "LButton", "U"

A_Clipboard := ""
Send "^c"
ClipWait(0.25)

If A_Clipboard ~= "\t\d{5}-\d{4}-\d{2}\t"
    NDCArray.Push(SubStr(A_Clipboard,RegExMatch(A_Clipboard,"\t\d{5}-\d{4}-\d{2}\t")+1,13))

SetTimer ProductSelected, 100

ProductSelected(){
    If !WinExist("Product Selection")
        {
        SetTimer ProductSelected, 0
    ; Work backwards through array to find last non-empty value
        While NDCArray.Has(-1)
            {
            If NDCArray.Get(-1) = ""
                NDCArray.RemoveAt(-1)
            Else
                break
            }
; Use that value to query RxItem
If NDCArray.Has(-1)
    {
    ProdID := NDCArray.Get(-1)
    ; Reset static variables to prepare for next rx, using timer to delay restore
    ; Without timer, penultimate run resets NewRx prematurely, causing PriorClip to store wrong value
            RxItem(ProdID,"NDC")
            }
        }
    Return
    }
}
*/

/**************************************************************** Eligibility Lookup ****************************************************************
; This hotkey runs a query to check for CCN authorizations and SC information.
; Currently unable to extract DoD ID from application within workflow.
; Patient name +/- DOB can be extracted from some windows and used to populate a default ID.
; Results are returned via GUI window.
****************************************************************************************************************************************************/

#HotIf WinExist("ahk_group Elig")
!+p::
; Named function to allow hotkey to be called from GUIMain.
EligHK(hk)
{
If !LSVAccess
    Exit

If A_ThisHotkey = "!+p"
    KeyWait "Alt", "T2"

If WinExist("Pharmacy Eligibility")
    {
    WinActivate("Pharmacy Eligibility")
    AdjustOverlap("Pharmacy Eligibility Overview","PharmNet: Retail Medication Manager")
    }
Else if WinActive("Person Search - \\\\Remote")
    {
    WinGetPos(&X,&Y,,,"A")
    OpenWhere := "x" . X + 275 . " y" . Y + 130 . " w250 h100"
    Eligibility(WinGetID("A"),OpenWhere)
    }
Else if WinExist("Person Search - \\\\Remote")
    {
    WinGetPos(&X,&Y,,,"Person Search - \\\\Remote")
    OpenWhere := "x" . X + 275 . " y" . Y + 130 . " w250 h100"
    Eligibility(WinGetID("Person Search - \\\\Remote"),OpenWhere)
    }

; Attempt to extract Patient name +/- DOB from window title to search for DoD ID value to default.
Else if WinActive("i)( - Pharm|for Patient:).*- \\\\Remote")
    {
    MMRPatient(&PatientName, &DOB, &Query)
    PatientSearch := LSVQuery(Query,0,5,3)

    DefID := ""

    If PatientSearch
        {
        PatientArray := StrSplit(RegExReplace(PatientSearch,"\R"),"|")
; String result often results in additional blank row; check if exists and delete if yes.
; Preference would be to figure out why blank row is created and prevent but this works for now.
        If !PatientArray.Get(-1)
            PatientArray.RemoveAt(-1)

        If PatientArray.Length = 1
            {
            PatientRow := StrSplit(PatientArray[1],"`t")
            DefID := PatientRow.Has(2) ? PatientRow[2] : ""
            }
        }
    Eligibility(,,DefID)
    }
Eligibility()
}
#HotIf

; Separated hotkey rather than stacking because the Eligibility function will close any existing window related to this hotkey
#HotIf WinActive("\QPerson Search - \\Remote\E")
!LButton::
{
KeyWait "Alt", "T2"
WinGetPos(&X,&Y,,,"A")
OpenWhere := "x" . X + 275 . " y" . Y + 130 . " w250 h100"
Eligibility(WinGetID("A"),OpenWhere)
}
#HotIf

/************************************************************ Alt + Shift + U (Unclaim) ***********************************************************************
; This hotkey is used to process unclaim actions
; Updated 11/22/2023 to include ability to check for DoD HP and use modify action in CM to then call CheckHP to remove/replace DoD plan
;   Option is controlled by GUIUnclaim
;   Modify ends with resubmit so list refreshes; this prevents "mixing" batches and a message box is expected to alert if unclaimed non-dummy rejects need submission
; GUIUnclaim cannot be separated from hotkey unless local variables are replaced with global (maybe static too) - no reason to separate so nesting
**************************************************************************************************************************************************************/

#HotIf WinActive("PharmNet: Claims Monitor")
!+u::
; Named function to allow hotkey to be called from GUIMain.
Unclaim(hk)
{
If A_ThisHotkey = "!+u"
    KeyWait "Alt", "T2"

yPos := 200			            ; Expected to be row 1 of CM since CoordMode is set to Window this should be reliable across users and computers, etc.
Unsubmitted := ""	            ; Used to track pending unclaim Rxs.
Skipped := ""	                ; Used to track Rxs skipped due to multiple claims in worklist.
UnsubFlag := 0                  ; Used to track if user has been alerted to unsubmitted claims.

; Call GUI with ByRef parameters so that we can pass values back to the hotkey.
GUIUnclaim(&RowCount, &ModifyHP, &CheckRTS, &Reversal)

; If GUI is closed or user enters value other than positive integer, we'll exit.
If !IsSet(RowCount) OR RowCount <= 0
    Exit

; Temporarily block input that could disrupt script.
BlockControl(2000)

; Maximize window to semi-standardize amount of rows available.
; And define end of list position
If !WinActive("\QPharmNet: Claims Monitor - \\Remote\E")
    WinActivate("\QPharmNet: Claims Monitor - \\Remote\E")
If !WinGetMinMax("\QPharmNet: Claims Monitor - \\Remote\E")
    MaxWin("\QPharmNet: Claims Monitor - \\Remote\E")
WinGetPos(,,,&H,"\QPharmNet: Claims Monitor - \\Remote\E")
yPosMax := H - 66

; This is the main loop that checks the active row and acts accordingly
Loop RowCount
	{
; Display tool tip and activate Claims Monitor window
    ToolTip "Please wait, script is working row " A_Index " of " RowCount ".", 881, 32
    WinActivate("\QPharmNet: Claims Monitor - \\Remote\E")

; Removal of DoD HP ends with submit, which refreshes list
; This loop is here to ensure we are working a new row
    Loop 40
        {
        ThisClaim := CMRow(yPos)
        ; If claim object doesn't exist or the claim is 'Inprocess', try again
        If !IsObject(ThisClaim) OR ThisClaim["Status"] = "Inprocess"
            {
            Sleep 250
            Continue
            }

        ; If this claim object exists and lastclaim string doesn't, we're on first row
        Else if IsObject(ThisClaim) AND !IsSet(LastClaim)
            Break

        ; CMRow keeps A_Clipboard intact so we can compare the full strings
        Else if A_Clipboard != LastClaim
            Break

        ; Sleep is generally not necessary when scrolling the list, but lag is noticeable and annoying
        ; So only sleep if we've tried at least once already
        Else if A_Index > 1
            Sleep 250
        }

    ; If the row has a line break in the sig, the row height is bigger, so we will try to adjust
    ; If map key exists, check for presence of line break characters
    ; Tried storing Match.Count from RegExMatch for > 1 line break but haven't gotten this to work and it's low priority
    If ThisClaim.Has("Medication") AND RegExMatch(ThisClaim["Medication"],"\n.")
        yInc := 30
    Else yInc := 18

    ; Check that we can proceed
    ;  - Multiple claims can lead to rows becoming 'unclaimed' earlier than expected
    ;  - If selected row has not changed, then try moving to next row
    ; If incrementing yPos would exceeed max, send down instead and give app half second to catch up
    ; Else increment yPos to next row
    If (ThisClaim["Transaction"] = "Reverse" AND (!Reversal OR !ThisClaim["HxFill"])) OR (IsSet(LastClaim) AND ThisClaim = LastClaim)
        {
        If yPos + yInc > yPosMax
            {
            Send "{Down}"
            Sleep 500
            }
        Else yPos += yInc
        Continue
        }
    Else if ThisClaim["Transaction"] = "Reverse" AND Reversal AND ThisClaim["HxFill"]
        {
        CMReverse(ThisClaim, yPos)
        If yPos + yInc > yPosMax
            {
            Send "{Down}"
            Sleep 500
            }
        Else yPos += yInc
        Continue
        }
    ; This row was unclaimed earlier in the script so move to next row
    Else if ThisClaim["Status"] = "Unclaimed" AND InStr(Unsubmitted,ThisClaim["RxNumber"])
        {
        If yPos + yInc > yPosMax
            {
            Send "{Down}"
            Sleep 500
            }
        Else yPos += yInc
        OutOfOrder := 1
        Continue
        }
    ; We've already logged an unclaim for this Rx, so add it to the skipped list and move to next row
    Else if InStr(Unsubmitted,ThisClaim["RxNumber"])
        {
        If yPos + yInc > yPosMax
            {
            Send "{Down}"
            Sleep 500
            }
        Else yPos += yInc
        If NOT InStr(Skipped,ThisClaim["RxNumber"])
            Skipped .= ThisClaim["RxNumber"] "`n"
        Continue
        }

    ; We're done with the new claim check so store copied row before next possible use of clipboard.
    LastClaim := A_Clipboard

    ; Non-dummy HP unclaims will be lost if unsubmitted and modify triggers refresh.
    If ThisClaim["DummyHPFlag"] AND Unsubmitted AND UnsubFlag = 0
        {
        UnsubFlag := 1
        UnsubOK := MsgBox("Unsubmitted claims will be lost. Please submit pending unclaim actions in Claims Monitor, then click OK to continue script. Or click Cancel to exit script.","Unsubmitted Claims","OC")

        If UnsubOK != "OK"
            Exit
        Else yPos := 200, UnSubFlag := 0

        BlockControl(2000)
        }

    ; Use CMReject and check return value.
    Resolved := CMReject(ThisClaim, yPos, ModifyHP, CheckRTS)

    ; If returns false or this row was NOT a dummy.
    If !Resolved OR ThisClaim["DummyHPFlag"] = 0
        {
        ; Update yPos relative to window.
        If yPos + yInc > yPosMax
            {
            Send "{Down}"
            Sleep 500
            }
        Else yPos += yInc

        ; And if NOT dummy, update tracking variables
        If ThisClaim["DummyHPFlag"] = 0 AND NOT InStr(Unsubmitted,ThisClaim["RxNumber"])
            Unsubmitted .= ThisClaim["RxNumber"] "`n"
        }
	}

; Disable tool tip, release blocks, notify user
ToolTip
BlockInput "Off"
BlockInput "MouseMove" "Off"
PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\CLICK_Nice.mp3")
If IsSet(OutOfOrder)
    MsgBox("Previously unclaimed rows were identified during processing that may fail to unclaim during submission. If needed, try sorting the Claims Monitor list by Date/Time (descending) then run the script again.","Unclaim Script Complete but...","4096 T10") 
Else if Skipped
    MsgBox("The Rxs below had multiple pending claims and at least one claim was skipped. Run script again to work these claims:`n`n" Skipped,"Let's Do It Again!","4096 T7")
Else
    MsgBox("All done!","Unclaim Script Complete","4096 T4")
Exit

; GUI to receive variables for unclaim script.

GUIUnclaim(*){
HelpMenu := Menu()
HelpMenu.Add "&OCRAP-VA Help", (*) => Run("https://dvagov.sharepoint.com/sites/PharmacyCernerEHRCommunityofPractice/SitePages/Oracle-Cerner-Repository-for-Automating-Pharmacy-VA.aspx?csf=1&web=1&share=EZYFbCtdCr9Bvseh7gKT8tMB4MMSjq_qW1N5PuvnfEAATw&e=o2SoTo&cid=cfb91149-243b-4ac4-b714-d4e64239a0ad#description-14")
Menus := MenuBar()
Menus.Add "&Help", HelpMenu
UnclaimGUI := Gui()
UnclaimGUI.Opt("+AlwaysOnTop")
UnclaimGUI.MenuBar := Menus
UnclaimGUI.Title := "Alt+Shift+U (Unclaim) Hotkey"
UnclaimGUI.AddText("", "This script will attempt to resolve claim rejects in Claims Monitor.`n`nHow many rows would you like to attempt?")
UnclaimGUI.AddEdit("r1 vRowCount w24 x275 y28")
If CheckHPOn = 1
    {
    UnclaimGUI.AddText("xm","Check to remove DoD Health Plans via Modify:")
    UnclaimGUI.AddCheckBox("vModify x275 yp Checked")
    }
UnclaimGUI.AddText("xm","Check to screen RTS next fill date before unclaim:")
UnclaimGUI.AddCheckBox("vRTS x275 yp Checked")
UnclaimGUI.AddText("xm","Reverse unassociated/historical bills:")
UnclaimGUI.AddCheckBox("vReversals x275 yp")

UnclaimGUI.AddButton("Section w80 x120 default", "&OK").OnEvent("Click", Submit)
UnclaimGUI.OnEvent("Close", (*) => UnclaimGUI.Destroy())
UnclaimGUI.OnEvent("Escape", (*) => UnclaimGUI.Destroy())
UnclaimGUI.Show "w320"
; Wait for GUI window to close before returning call.
WinWaitClose(UnclaimGUI.Hwnd)

; Function to save GUI input.
Submit(*){
Saved := UnclaimGUI.Submit(1)
; If RowCount value is non-integer, set it to zero so script will exit.
IsInteger(Saved.RowCount) ? RowCount := Saved.RowCount : RowCount := 0

ModifyHP := CheckHPOn = 1 ? Saved.Modify : 2
CheckRTS := Saved.RTS
Reversal := Saved.Reversals
WinActivate("PharmNet: Claims Monitor")
Return
}
}
}
#HotIf

/************************************************************ Alt + Shift + O (Open) *********************************************************************
; This hotkey is used to open Citrix Storefront.
*************************************************************************************************************************************************************/

!+o::
{
Run ("https://login.valehr.cernerworks.ehr.gov/Citrix/USVADCweb/#")
; Specify below the path to the GIF file to animate (local files are allowed too):
Try
    {
    /* Randomizer!
    ; Calculate browser position (centered on mouse).
    browserWidth := 500
    browserHeight := 350
    browserX := MonitorMiddle - (browserWidth / 2)
    browserY := 200
    ;Random gif generator: https://github.com/Onyx-Nostalgia/gifroz
    Run "msedge.exe --app=https://gifroz.vercel.app/?q=red%20panda --window-position=" browserX "," browserY
    If WinWait("i)gifroz\.vercel\.app \(\d+.\d+\)",,30)
        {
        ; Dim := StrSplit(RegExReplace(WinGetTitle("i)gifroz\.vercel\.app "),"i)gifroz\.vercel\.app "),"×",,2)
        ; WinMove(, , Dim[1], Dim[2],"i)gifroz\.vercel\.app \(\d+.\d+\)")
        WinSetAlwaysOnTop(1,"i)gifroz\.vercel\.app")
        Sleep 4000
        If WinExist("i)gifroz\.vercel\.app")
            WinClose("i)gifroz\.vercel\.app")
        }
    */ ;Curated
    LinkGIF := [ ;List URLs.
        {url: "https://media3.giphy.com/media/v1.Y2lkPTc5MGI3NjExMm1laWttcmRzeDQ3MHlwbnA3MXN1M3VsNHhweHEyMnNqbnpwbG4xMCZlcD12MV9pbnRlcm5hbF9naWZfYnlfaWQmY3Q9Zw/KEf7gXqvQ8B3SWnUid/giphy.gif", width: 480, height: 480},
        {url: "https://media4.giphy.com/media/v1.Y2lkPTc5MGI3NjExZjdqZDVqNGJhOWN2anpmcTA2bG03YmgwNzQwZWo3Y2YzMmdjdDA0aiZlcD12MV9pbnRlcm5hbF9naWZfYnlfaWQmY3Q9Zw/Umb1ju55UMAIE/giphy.gif", width: 397, height: 222},
        {url: "https://media2.giphy.com/media/v1.Y2lkPTc5MGI3NjExYmZrZnlkejFjYzk3dWl5cGZxY3l1ZXBzaG5haXZzODJrbmhlZnF2NiZlcD12MV9pbnRlcm5hbF9naWZfYnlfaWQmY3Q9Zw/3ov9jIMfsR1wRSeO9W/giphy.gif", width: 480, height: 386},
        {url: "https://media0.giphy.com/media/v1.Y2lkPTc5MGI3NjExMTlmZ2lqZW5kcmRxNGU4OTgwaXhhdWpqd2w2dmd2a2E2Z2lncXVmMCZlcD12MV9pbnRlcm5hbF9naWZfYnlfaWQmY3Q9Zw/IThjAlJnD9WNO/giphy.gif", width: 245, height: 184},
        {url: "https://media2.giphy.com/media/v1.Y2lkPTc5MGI3NjExaXAzeHcxZnQ2dW0yNnVuNmJseDZwYWNodTF2c2Q3NWdyYXpneGRxbiZlcD12MV9pbnRlcm5hbF9naWZfYnlfaWQmY3Q9Zw/NZmrvsfxpbTSU/giphy.gif", width: 480, height: 222},
        {url: "https://media0.giphy.com/media/v1.Y2lkPTc5MGI3NjExMTV3eXAyaGoxMXFsa3A1anc0NGwzZDQxZWc4M2tkb3YyZGZoeG5ycyZlcD12MV9pbnRlcm5hbF9naWZfYnlfaWQmY3Q9Zw/Q0LdqbADEDDmE/giphy.gif", width: 480, height: 244},
        {url: "https://media1.giphy.com/media/v1.Y2lkPTc5MGI3NjExbzU4ZHkydjIwZm42cnFyazQ5dHgxbTJpd3VwaTU4NGFpZm8ycDYxNiZlcD12MV9pbnRlcm5hbF9naWZfYnlfaWQmY3Q9Zw/BVStb13YiR5Qs/giphy.gif", width: 500, height: 353},
        {url: "https://media4.giphy.com/media/v1.Y2lkPTc5MGI3NjExbm16aTR0a3k2aXlneHprbzZlaWNqbHdvYzB3NzI0ajYzaDF1NTlqZiZlcD12MV9pbnRlcm5hbF9naWZfYnlfaWQmY3Q9Zw/QtwIscCaoWkxiD1mjM/giphy.gif", width: 431, height: 511},
        {url: "https://media1.giphy.com/media/v1.Y2lkPTc5MGI3NjExbTBlcWluMXU4OXdjbWs5cno5aXdoYmc0aTk3NDhva3llNWkwamxlMiZlcD12MV9pbnRlcm5hbF9naWZfYnlfaWQmY3Q9Zw/LHxKBH8mIilmUeZRUO/giphy.gif", width: 480, height: 480},
        {url: "https://media2.giphy.com/media/v1.Y2lkPTc5MGI3NjExMGM3aHl6eXJrcGh3ajFpbHYwNm4xYWd5YTJyczIwazVieDJpYnZmeCZlcD12MV9pbnRlcm5hbF9naWZfYnlfaWQmY3Q9Zw/3og0IP4nKFrVWeB9M4/giphy.gif", width: 480, height: 270},
        {url: "https://media3.giphy.com/media/v1.Y2lkPTc5MGI3NjExNjl2cHBidDU0MmFqeHFzOWhlNWl6ZWQ1aDZ4cGVwMTE5ZmxwZWFpMSZlcD12MV9pbnRlcm5hbF9naWZfYnlfaWQmY3Q9Zw/Cmr1OMJ2FN0B2/giphy.gif", width: 500, height: 500},
        {url: "https://media1.giphy.com/media/v1.Y2lkPTc5MGI3NjExaWxnZjUxeHE1eHhqYm5kMWppOGM3ZzIzMjZoaHJzbmo4emIyczhqcSZlcD12MV9pbnRlcm5hbF9naWZfYnlfaWQmY3Q9Zw/kW8mnYSNkUYKc/giphy.gif", width: 480, height: 362},
        {url: "https://media0.giphy.com/media/v1.Y2lkPTc5MGI3NjExenc2dXNxY3NtMjhncXdrdTNraml3MXQ1czNhNWpmYnVkM3hyNnVndiZlcD12MV9pbnRlcm5hbF9naWZfYnlfaWQmY3Q9Zw/brsEO1JayBVja/giphy.gif", width: 500, height: 278},
        {url: "https://media0.giphy.com/media/v1.Y2lkPTc5MGI3NjExN3RveG9lM2EyaDJrNTJ5eWg3NHZsOTRwdDNranZxZTFtc205cnl3aSZlcD12MV9pbnRlcm5hbF9naWZfYnlfaWQmY3Q9Zw/2rAKTgJIQe1buYU1R5/giphy.gif", width: 380, height: 304},
        {url: "https://media1.giphy.com/media/v1.Y2lkPTc5MGI3NjExNjA2cHdjZ3NrYWczMmt3b3hibjNhdWF1amszMDAxbWFndm5jYjBuZyZlcD12MV9pbnRlcm5hbF9naWZfYnlfaWQmY3Q9Zw/2f7RQiiWMJc40/giphy.gif", width: 343, height: 285}
    ]
    RandomLink := Random(1, LinkGIF.Length) ;Spin the wheel!
    GIFURL := LinkGIF[RandomLink] ;Choose
    URLGIF := GIFURL.url
    w := GIFURL.width + 45
    h := GIFURL.height + 45
    ;Display
    GIFGUI := Gui("+AlwaysOnTop","Hello and Welcome")
    wb := GIFGUI.Add("ActiveX", "w" w " h" h, "Shell.Explorer").Value
    wb.Navigate(URLGIF)
    GIFGUI.Show("w" w " h" h)
    Sleep 4000 ;Timer
    If WinExist("Hello and Welcome")
        GIFGUI.Destroy
    }
Catch
    MsgBox("Good luck today!","Welcome","4096 T4")
}

/************************************************************ Alt + Shift + F (Failures) *********************************************************************
; This hotkey uses various pathways to resolve failures in the Work Queue Monitor
; The GUI will allow the user to select a pathway and then run the appropriate function
*************************************************************************************************************************************************************/

#HotIf WinActive("Work Queue Monitor")
; Alt+Shift+F (Failures)
!+f::
; Named function hotkey so that this can be called from GUIMain.
Failures(hk, Pathway := ""){
; Call Failure GUI - return value will determine next step.
If Pathway = ""
    GUIFailWelcome(&ScriptPath)

Else ScriptPath := Pathway

; If user closes GUI or the ScriptPath variable is not set, so we exit the script.
If !IsSet(ScriptPath)
	Exit

; This pathway will scroll a view and attempt to resolve via WQMFail function.
Else if ScriptPath = "View"
    FailuresBatchView()

; This pathway will complete resuspend actions initiated by a user.
Else if ScriptPath = "Resuspend"
    FailuresBatchResuspend()

; This pathway is used to query for user inactivated prescriptions with a 'pre-transmission' failures.
Else if ScriptPath = "Query"
	FailuresBatchQuery()

; This pathway uses a query to screen for 'narrow' Suspend Until Date validation failures (will bypass Controlled Substances).
Else if ScriptPath = "SUD"
    FailuresBatchSUD()

; This pathway is used to resuspend dispenses from Window to CMOP.
Else if ScriptPath = "Window"
    FailuresBatchWindow()

; This pathway uses a query to look for dispenses suspended for local mail that may be CMOP eligible.
Else if ScriptPath = "Local Mail"
    FailuresBatchLocalMail()

; ScriptPath variable does not match a case above.
Else Exit

; GUI to handle user input.
GUIFailWelcome(&ScriptPath)
{
FileMenu := Menu()
FileMenu.Add "E&xit", (*) => FailGUI.Destroy()
HelpMenu := Menu()
HelpMenu.Add "&OCRAP-VA Help", (*) => Run("https://dvagov.sharepoint.com/sites/PharmacyCernerEHRCommunityofPractice/SitePages/Oracle-Cerner-Repository-for-Automating-Pharmacy-VA.aspx?csf=1&web=1&share=EZYFbCtdCr9Bvseh7gKT8tMB4MMSjq_qW1N5PuvnfEAATw&e=o2SoTo&cid=cfb91149-243b-4ac4-b714-d4e64239a0ad#use-a-specific-view")
Menus := MenuBar()
Menus.Add "&File", FileMenu
Menus.Add "&Help", HelpMenu
FailGUI := Gui()
FailGUI.Opt("+AlwaysOnTop")
FailGUI.MenuBar := Menus
DocLinkText := '- Suggested view: restricted to qualifying failures <a href="https://dvagov.sharepoint.com/sites/PharmacyCernerEHRCommunityofPractice/SitePages/Oracle-Cerner-Repository-for-Automating-Pharmacy-VA.aspx?csf=1&web=1&share=EZYFbCtdCr9Bvseh7gKT8tMB4MMSjq_qW1N5PuvnfEAATw&e=o2SoTo&cid=cfb91149-243b-4ac4-b714-d4e64239a0ad#use-a-specific-view">(see documentation)</a>'
QueryLinkText := format('- Suggested view: Pharmacy + Validation check failures <a href="https://vaww.pbi.cdw.va.gov/pbi_rs/report/FRE/D05_VISN20/SSRS/Production%20Reports/Cerner%20Work%20Queue%20Monitor%20Failures?Sta6a={1}">(see query)</a>',IsSet(Sta3n) ? Sta3n : "")
FailGUI.SetFont("s9.5")
FailGUI.Title := "Alt+Shift+F (Failures) Hotkey"
FailGUI.AddText("Section w480","Click here to have the script start at row 1 and attempt to resolve qualifying failures`n- Suggested use: daily")
FailGUI.AddLink("y+0", DocLinkText)
FailGUI.AddText("w480","Click here to have the script complete resuspend actions`n- Suggested use: CMOP stock kickbacks to local, CMOP batch fail to CMOP`n- Suggested view: Pretty flexible really")
FailGUI.AddText("w480","Click here to have the script resuspend pending window refills`n- Suggested use: daily before running local mail batches`n- Suggested view: Window Suspended T+0")
If LSVAccess
    {
    FailGUI.AddText("w480","Click here to have the script delete pending refills of orders inactivated by a user`n- Suggested use: ad hoc as indicated by query results")
    FailGUI.AddLink("y+0", QueryLinkText)
    FailGUI.AddText("w480","Click here to have the script identify SUD failures with wiggle room`n- Suggested use: ad hoc as indicated by SUD fail volume")
    FailGUI.AddText("w480","Click here to have the script resuspend future local mail refills to CMOP`n- Suggested use: ad hoc`n- Suggested view: Local Suspended T+15-30")
    }

FailGUI.AddButton("ys+8","View").OnEvent("Click", EightfoldPath)
FailGUI.AddButton("y+35","Resuspend").OnEvent("Click", EightfoldPath)
FailGUI.AddButton("y+35","Window").OnEvent("Click", EightfoldPath)
If LSVAccess
    {
    FailGUI.AddButton("y+35","Query").OnEvent("Click", EightfoldPath)
    FailGUI.AddButton("y+35","SUD").OnEvent("Click", EightfoldPath)
    FailGUI.AddButton("y+35","Local Mail").OnEvent("Click", EightfoldPath)
    }
FailGUI.AddText("xm y+30","Please make sure you have an appropriate view selected and that you have read,`nand understand, the documentation for this script prior to use.")

FailGUI.OnEvent("Close", (*) => FailGUI.Destroy())
FailGUI.OnEvent("Escape", (*) => FailGUI.Destroy())
FailGUI.Show "w660"

; Wait for GUI window to close before returning call.
WinWaitClose(FailGUI.Hwnd)

; Update ScriptPath variable and do next thing based on GUI input.
EightfoldPath(GuiCtrlObj, Info){
ScriptPath := GuiCtrlObj.Text
FailGUI.Destroy()
Return
}
}

/**************************************************** Failures Batch Actions **************************************************
; There isn't necessarily an advantage to having these as functions but the script flow was becoming difficult to follow.
*******************************************************************************************************************************/

FailuresBatchView(){
RowCount := InputBox("You selected View. The script will scroll the loaded view to evaluate and attempt to resolve failures in view.`n`nHow many rows would you like to attempt?","Alt+Shift+F (Failures) Hotkey - View","w400 h150")

If RowCount.Result != "OK"
    Exit

If !RegExMatch(RowCount.Value, "(\d+)", &RC)
    {
    MsgBox("Must enter a whole number.", "Missing Row Count", "4096 T3")
    Exit
    }

If !WinActive("Work Queue Monitor - \\\\Remote")
    WinActivate("Work Queue Monitor - \\\\Remote")
If !WinGetMinMax("Work Queue Monitor - \\\\Remote")
    MaxWin("Work Queue Monitor - \\\\Remote")

; Define initial y-position (row 1)
; And placeholder for last Rx number to check that list has refreshed/we're working new row
yPosMouse := 153
LastRx := "0"

WinActivate("Work Queue Monitor")

; Loop as many times as user defined.
Loop RC[1]
    {
; Display tool tip so user knows script is running.
    ToolTip "Please wait, script is working row " A_Index " of " RC[1] ".", 875, 32
    WinActivate("Work Queue Monitor - \\\\Remote")

    WQM := WQMRow(yPosMouse)

    ; If RxNumber = LastRx
    If WQM["RxNumber"] = LastRx
        {
        Sleep 500
        A_Index -= 1
        Continue
        }
    Else if !WQM.Has("FailureDetail") OR WQM["FailureDetail"] = ""
        {
        yPosMouse += 18
        Continue
        }
/*
    Else if WQM["RequestMethod"] = "Internal" AND WQM["FailureDetail"] = "The product has been inactivated" ;B13 appears to have removed WQM's ability to handle these itself.
        {
        yPosMouse += 18
        Tooltip "Conversion required. Please trigger with Ctrl+Lclick when this is done. Notify your local drug file manager if not mapped yet.",,,2
        SetTimer () => ToolTip(,,,2), -3000
        Continue
        }
*/
    WQMFail(WQM, &yPosMouse)

    LastRx := WQM["RxNumber"]
    }
ToolTip
}

/************************************************* Resuspend pathway **********************************************
This pathway completes resuspend actions that were previously assigned by the user (or another script).
Basically this is a way to clear kickbacks that just need to be re-routed (e.g. CMOP stock => local;
CMOP 020: Database => CMOP ) in an a la carte manner (or dump everything if that's your preference)
******************************************************************************************************************/

FailuresBatchResuspend(){
GUIResuspend(&Destination,&CancelReason)

If !IsSet(Destination) OR !IsSet(CancelReason)
    Exit

WinGetPos , , , &h, "Work Queue Monitor"

ToolTip "Script is executing.", 875, 32
Loop
    {
    Try ReadyLoadRecords := WQMResuspend(Destination,CancelReason,2)
    Catch Error as err
        {
        If A_Index > 1
            {
            LogError(err)
            Tooltip
            PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\sad trombone.mp3")
            MsgBox(err.Message ' Exiting script.',"Unable to Activate Window","T5 4096")
            WinWaitClose("Unable to Activate Window")
            PlaySound()
            Exit
            }
        Else
            {
            PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\60 Minutes in 4s.mp3")
            OKGo := MsgBox("Waiting for Refill or Bill-Only (cancel fill) window to load. If you have not yet hit submit, please do so now, then click OK to continue. Click cancel to exit.","Window Shopping","OC 4096")
            If OKGo != "OK"
                Exit
            WinMatch("(Refill|Bill-Only) for Patient",2)
            }
        }
    If WinWaitActive("Work Queue Monitor",,3)
        {
        If ReadyLoadRecords
            Break
        If ColorFilter = 2
            {
            Send "#^c" ;Disable color filter.
            Sleep 500
            }
        If ImageSearch(&X, &Y, 0, h-30, 50, h, "*75 HBITMAP:*" ImageReady)
            OR ImageSearch(&X, &Y, 0, h-30, 50, h, "*75 HBITMAP:*" ImageLoading)
            OR ImageSearch(&X, &Y, 0, h-30, 50, h, "*75 HBITMAP:*" ImageNoRecords)
            {
            If ColorFilter = 2
                Send "#^c" ;Enable color filter.
            Break
            }
        If ColorFilter = 2
            Send "#^c" ;Enable color filter.
        }
    }
BlockInput "Off"
BlockInput "MouseMove" "Off"
ToolTip
PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\CLICK_Nice.mp3")
MsgBox("Script is complete",Format("Alt+Shift+F (Failures) Hotkey - Resuspend {1}",Destination),"4096 T10")
Exit

GUIResuspend(&Destination,&Reason,*)
{
RSGUI := Gui()
RSGUI.Opt("+AlwaysOnTop")
RSGUI.Title := "Alt+Shift+F - Resuspend"

;Instructions
RSGUI.AddText("w320", "You selected Resuspend. Please note: This is a 'simple' resuspend, only the routing option and health plan (if update health plan is enabled) may be updated. Intervention may be required to address failures, alerts, etc.")

; Bold part.
RSGUI.SetFont("Bold")
RSGUI.AddText("w320", "Before clicking 'Go!' below, please populate the resuspend action for any rows to be resuspended and ")

; Bold + underline part.
RSGUI.SetFont("Bold Underline")
RSGUI.AddText("y+0", "submit in WQM!")

; Reset to normal font.
RSGUI.SetFont("Norm")
RSGUI.AddText("w320", "The script will then take over to complete the resuspend sequence.`n`nIf resolving a CMOP batch failure, select 'No Change' as the resuspend option below.`n")

RSGUI.AddText("xm w140","Resuspend Routing Option:")
RSGUI.AddDropDownList("x+10 yp-2 w150 BackgroundWhite vNewRoute",["CMOP","Local","No Change"])
RSGUI.AddText("xm w140","Cancel Reason:")
RSGUI.AddDropDownList("x+10 yp-2 w150 BackgroundWhite vReason",["Pharmacist Judgment","Pharmacy Out of Stock","CMOP Cancellation","Formulary Substitution","Medication Recall","Patient Request"])
RSGUI.AddButton("Section w60 x140 Default", "&Go!").OnEvent("Click", Submit)
RSGUI.OnEvent("Close", (*) => RSGUI.Destroy())
RSGUI.OnEvent("Escape", (*) => RSGUI.Destroy())
RSGUI.Show

; Wait for GUI window to close before returning call.
WinWaitClose(RSGUI.Hwnd)

; Function to save GUI input.
Submit(*){
Saved := RSGUI.Submit(1)
Destination := Saved.NewRoute, Reason := Saved.Reason
Return
}
}
}

FailuresBatchQuery(){
; station value as defined in GUI Main, will use as default for Input Box
global Sta3n
Site := InputBox("You selected query. Upon clicking OK, a query will run then the script will scroll the loaded view to find identified prescription(s). If found, a delete action will be applied.`n`nPlease enter your station number below:","Alt+Shift+F (Failures) Hotkey - Query","w300",Sta3n)

If Site.Result != "OK" OR !RegExMatch(Site.Value, "^\d{3}[A-Za-z]?[A-Za-z]?$")
    Exit
Else
    Sta6a := Site.Value

; Display tooltip.
ToolTip "Running query for " Sta6a "...", 875, 32

; Get a list of RxNumbers with a dispense to be deleted and upper limit for number of cancel-fills to be performed.
DeleteRx := LSVQuery(Format("
    (
    SELECT String_Agg(RxNumber,', ')
        , SUM(InternalRequest)
        , MAX(RxNumber)
        , MIN(RxNumber)
    FROM LSV.D05_VISN20INFORM.AHK_WQM_Failure
    WHERE Sta6a = '{1}'
    GROUP BY Sta6a
    )",Sta6a),1)

; If query has no results, disable tooltip, message user, exit
If !DeleteRx
    {
    ToolTip
    PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\CLICK_Nice.mp3")
    MsgBox("No prescriptions found via query. Exiting script.","Nothing To Do Here","4096 T4")
    Exit
    }

; Parse query results
RxArray := StrSplit(DeleteRx,"`t")
RxNumbers := RxArray[1]
CancelCount := Integer(RegExReplace(RxArray[2],"\D"))
MaxRxNumber := Integer(RegExReplace(RxArray[3],'\D'))
MinRxNumber := Integer(RegExReplace(RxArray[4],'\D'))

; Display tool tip
ToolTip "Searching WQM for user inactivated prescriptions`nRx Range: " MinRxNumber " - " MaxRxNumber, 875, 32

; Counter that tracks if actions are applied so we don't submit zero actions
DelCount := 0

; User may have SSRS report loaded so we need to make sure we activate/maximize the WQM application
If !WinActive("Work Queue Monitor - \\\\Remote")
    WinActivate("Work Queue Monitor - \\\\Remote")
If !WinGetMinMax("Work Queue Monitor - \\\\Remote")
    MaxWin("Work Queue Monitor - \\\\Remote")
WinGetPos , , , &hWQM, "Work Queue Monitor - \\\\Remote"

If ColorFilter = 2
    {
    Send "#^c" ;Disable color filter.
    Sleep 500
    }
Loop 3
    {
    ImageSearch(&xHS, &yHS, 0, hWQM - 150, 40, hWQM - 80, "*75 HBITMAP:*" ImageWQMHorizontalScroll)
    Sleep 250
    } until xHS != "" AND yHS != ""
If ColorFilter = 2
    Send "#^c" ;Enable color filter.

; Table/list starts at yPos 144, each row is 18 pixels tall
LastFullRow := yHS = "" ? Floor((hWQM - 95 - 144) / 18) : Floor((yHS - 144) / 18)

; Try to find Rx Number column using ImageSearch.
If ColorFilter = 2
    {
    Send "#^c" ;Disable color filter.
    Sleep 500
    }
Loop 3
    {
    ImageSearch(&X, &Y, 0, 0, 1920, 175, "*75 HBITMAP:*" ImageRxNumHeader)
    Sleep 250
    } until X != "" AND Y != ""
If ColorFilter = 2
    Send "#^c" ;Enable color filter.

; If successful, offset to aim click else assume column 2
If X != "" AND Y != ""
    X += 30, Y += 10, Sorted := 1
If X = "" OR Y = ""
    X := 200, Y := 132, Sorted := 0

; Sort Rx Number column to support min/max checks
; Sorting places cursor in row 1 of clicked column
MouseClick "left", X, Y
Sleep 500

LastRx := "0", MaxCheck := 0, MinCheck := 0, Dir := ""

Loop
    {
; Increment position if we are not yet at bottom of visible list
; Initial mouse position is column header, so we need to adjust on first pass to align with selected row
    If A_Index <= LastFullRow
        Y += 18

    Loop 3
        {
        A_Clipboard := ""
        Send "^c"
        If Clipwait(1)
            Break
        }

    RxNumber := StrReplace(A_Clipboard,"`r`n")
    If RegExMatch(RxNumber,"\d\d\d\d-\d\d\d\d\d\d\d\d") = 0 AND DelCount > 0
        Break
    Else if RegExMatch(RxNumber,"\d\d\d\d-\d\d\d\d\d\d\d\d") = 0
        {
        ToolTip
        MsgBox Format("Copied value ({1}) may not be an Rx Number. Exiting.",RxNumber)
        Exit
        }
; If RxNumber doesn't change we're at the end of the list, so break this loop
    Else if LastRx = RxNumber
        Break
; If RxNumber exists in result set (Haystack,Needle) then apply the recommended action (delete)
    Else if InStr(RxNumbers,RxNumber) > 0
        {
        Send "{Home}"
        Send "d"
        Send "{Enter}"
        MouseClick "left", X, Y
        Sleep 400
; Increment counter when action applied
        DelCount += 1
        }
    Send "{Down}"
    Sleep 250

; The boundary values may not exist when the script is run, so also check direction
    If LastRx != "0" AND Dir = ""
        Dir := StrReplace(LastRx,"-") > StrReplace(RxNumber,"-") ? "Desc" : "Asc"
    LastRx := RxNumber

; If current value is above MaxRxValue or
; If we are now below now below min value and direction is asc (moving up)
; Update MaxCheck
    If MaxCheck = 0
        AND (StrReplace(RxNumber,'-') >= MaxRxNumber
        OR (MinCheck = 1
        AND Dir = "Desc"
        AND Sorted = 1))
        {
        MaxCheck := 1
        }
; If current value is less than MinRxValue or
; If we are now above max value and sort direction is desc (moving down)
; Update MinCheck
    If MinCheck = 0
        AND (StrReplace(RxNumber,'-') <= MinRxNumber
        OR (MaxCheck = 1
        AND Dir = "Asc"
        AND Sorted = 1))
        {
        MinCheck := 1
        }
    If MaxCheck = 1
        AND MinCheck = 1
        AND (StrReplace(RxNumber,'-') >= MaxRxNumber
        OR StrReplace(RxNumber,'-') <= MinRxNumber)
        {
        Break
        }
    }

; No actions were applied, disable tooltip, message user and exit
If DelCount = 0
    {
    ToolTip
    PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\CLICK_Nice.mp3")
    MsgBox("Appears no actions were taken. Exiting script.","Nothing More To Do Here","4096 T4")
    Exit
    }

; Once list has been checked and at least one action applied, alt+s to submit
Send "!s"

; At least one dispense identified via query is an "internal" refill
If CancelCount > 0
    {
    ToolTip DelCount " delete actions submitted", 875, 32, 1
    LoopCount := CancelCount < DelCount ? CancelCount : DelCount
    WinGetPos , , , &h, "Work Queue Monitor - \\\\Remote"
    Cancels := 0

    If ColorFilter = 2
        {
        Send "#^c" ;Disable color filter.
        Sleep 500
        }
    Loop LoopCount
        {
        ToolTip "Completing cancel fill action #" A_Index "`nUp to " LoopCount " cancel(s) may be performed", 875, 32
; Cancel-Fill window can be very slow to load so wait a bit before entering loop
        If !WinWaitActive("Cancel-Fill",,30)
            {
            PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\60 Minutes in 4s.mp3")
            OKGo := MsgBox("Waiting for the Cancel-Fill window to load.`n`nIf you would like the script to continue waiting, leave this message box open, then click OK when a 'Cancel-Fill' window appears.`n`nIf WQM is done loading, or you would like the script to stop now, click Cancel to exit.","Internal Refill Cancel","4096 OC T300")
            If OKGo != "OK"
                {
                ToolTip
                Exit
                }
            Try WinActivate("Cancel-Fill")
            }
        If WinWaitActive("Cancel-Fill",,15)
            {
; Navigate to reason, populate, accept, wait for window to close
            Send "{Tab 2}"
            Send "Prescriber Modified Order"
            Send "!o"
            WinWaitNotActive("Cancel-Fill")
            Cancels += 1
            Sleep 1000
            }
        } Until (Cancels = LoopCount)
            OR ImageSearch(&X, &Y, 0, h-30, 50, h, "*75 HBITMAP:*" ImageReady)
            OR ImageSearch(&X, &Y, 0, h-30, 50, h, "*75 HBITMAP:*" ImageLoading)
            OR ImageSearch(&X, &Y, 0, h-30, 50, h, "*75 HBITMAP:*" ImageNoRecords)
    If ColorFilter = 2
        Send "#^c" ;Enable color filter.
    }

; Disable tooltip, notify user, exit
ToolTip
PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\CLICK_Nice.mp3")
MsgBox("Script is complete","Alt+Shift+F (Failures) Hotkey - Query","4096 T10")
Exit
}

FailuresBatchSUD(){
If !WinActive("Work Queue Monitor - \\\\Remote")
    WinActivate("Work Queue Monitor - \\\\Remote")
If !WinGetMinMax("Work Queue Monitor - \\\\Remote")
    MaxWin("Work Queue Monitor - \\\\Remote")
WinGetPos , , , &hWQM, "Work Queue Monitor - \\\\Remote"

GUISUDFail(&MaxDays,&Sorted)

If !IsSet(MaxDays) OR MaxDays = 0
    Exit

ToolTip "Script will start 1 second after a left-click in WQM", 875, 32, 1

WinActivate("PharmNet: Work Queue Monitor")
KeyWait "LButton", "D"
MouseGetPos , &yPosMouse
Sleep 1000

If ColorFilter = 2
    {
    Send "#^c" ;Disable color filter.
    Sleep 500
    }
Loop 3
    {
    ImageSearch(&xHS, &yHS, 0, hWQM - 150, 40, hWQM - 80, "*75 HBITMAP:*" ImageWQMHorizontalScroll)
    Sleep 250
    } until xHS != "" AND yHS != ""
If ColorFilter = 2
    Send "#^c" ;Enable color filter.

; Click defines 'top of list', each row is 18 pixels tall
; 9 is a 'fudge factor', assuming user clicks roughly in the middle of the row
; Table/list starts at yPos 144, each row is 18 pixels tall
; Bottom of list is defined as Horizontal scroll bar Y (yHS) if found, else bottom of list is 95 from bottom of window
RowsToBottom := yHS = "" ? Floor((hWQM - 95 - yPosMouse + 9) / 18) : Floor((yHS - yPosMouse + 9) / 18)

ToolTip "Please wait, script is executing.", 875, 32

SUDmap := Map()

Loop
    {
    If A_Index = 1
        LastRx := "0"
    Else
        LastRx := WQM["RxNumber"]

; Increment position if we are not yet at bottom of visible list
; A user click sets initial position so we do not adjust on A_Index 1
    If A_Index <= RowsToBottom AND A_Index > 1
        yPosMouse += 18

    If A_Index > 1
        {
        Send "{Down}"
        Sleep 400
        }

    WQM := WQMRow(yPosMouse)

; If RxNumber doesn't change or we didn't extract one we're at the end of the list, break loop
    If LastRx = WQM["RxNumber"] OR !WQM["RxNumber"] OR WQM["Action"] = "Action"
        Break
; If not a 'pure' SUD failure, break or continue depending on sort criteria
    Else if WQM["FailureDetail"] != "Suspend Until Date is greater than the Rx order stop date" AND Sorted = "Failure Detail"
        Break
    Else if WQM["FailureDetail"] != "Suspend Until Date is greater than the Rx order stop date"
        Continue
; If CS, skip
    Else if WQM["RxNumber"] ~= "^\d{4}-[2-4]\d{7}" OR WQM["LegalStatus"] ~= "^[1-5]"
        Continue
; If we fail to extract Rx Number or SUD, skip
    Else if !WQM["SUD"]
        Continue

; If checks pass, run query
; Would like to store result as map value (RxNumber, new DPD) but Citrix WinTitle is truncated and we can't see Rx Number
; Appears DLL call may be an option but skipping for now - https://www.autohotkey.com/boards/viewtopic.php?t=9870
    Query := Format("
        (
        SELECT CONVERT(varchar
                    ,IIF(ISNULL(ProjectedStopDateTime,DATEADD(YY,3,GETDATE())) < ExpireDateTime
                        ,ProjectedStopDateTime
                        ,ExpireDateTime),101)
                , IIF(OrderStatus = 'ORDERED', 1, 0)
        FROM LSV.D05_VISN20INFORM.RxMill_Prescription
        WHERE RxNumber = '{1}'
            AND ExpireDateTime > CAST(GETDATE() AS DATETIME2(0))
            AND (ProjectedStopDateTime > CAST(GETDATE() AS DATETIME2(0))
            OR ProjectedStopDateTime IS NULL)
            AND DATEDIFF(DD
                    , CAST(
                        IIF(ISNULL(ProjectedStopDateTime,DATEADD(YY,3,GETDATE())) < ExpireDateTime
                            ,ProjectedStopDateTime
                            ,ExpireDateTime) AS DATE)
                    ,'{2}') < '{3}'
        )",WQM["RxNumber"],SubStr(WQM["SUD"],1,10),MaxDays)

    RxCheck := LSVQuery(Query,1,5,3)

    If RxCheck
    {
; Parse query results.
    SUDArray := StrSplit(RxCheck,"`t")
    EndDate := SUDArray[1]
    ActiveOrderFlag := Integer(RegExReplace(SUDArray[2],"\D"))

; If result is active order and less than/equal to defined max spread, mark for resuspend.
    If ActiveOrderFlag = 1
        {
    ;    SUDmap[WQM["RxNumber"]] := EndDate ; See inactive code below. May be more useful one day.
        SUDmap[Integer(SUDmap.Count + 1)] := EndDate
        Send "{Right}"
        Send "{Home}"
        Send "r"
        Send "{Enter}"
        Sleep 400
        }

; If result is not active order, skip row and notify user without interrupting script.
    Else if ActiveOrderFlag = 0
        {
        ToolTip "Order inactive. Run delete query hotkey after this.", , , 2
        SetTimer () => ToolTip(,,,2), -2000
        }
    }

; No result... or query timed out, but those are caught by LSVQuery function earlier.
    Else if !RxCheck
        {
        ToolTip "SUD > " MaxDays " days beyond Rx end date.", , , 2
        SetTimer () => ToolTip(,,,2), -2000
        }
    }

; Ordinally resuspend eligible orders.
If SUDmap.Count
    {
    Send "!s"
    Loop
        {
        WinWaitActive("Refill for Patient:")
        If CheckHPOn = 1
            RxCheckHP()
        If ColorFilter = 2
            {
            Send "#^c" ;Disable color filter.
            Sleep 500
            }
        Loop 15
            {
            Sleep 200
            } until ImageSearch(&xDT, &yDT, 0, 450, 450, 850, "*75 HBITMAP:*" ImageRxDateTime)
        If ColorFilter = 2
            Send "#^c" ;Enable color filter.
        If xDT != "" AND yDT != ""
            MouseClick "left", xDT + 20, yDT + 30
        If xDT = "" OR yDT = ""
            {
            OKGo := MsgBox("Script failed to find Dispense Date/Time field. If you would like the script to continue, click OK here then click into the Dispense Date/Time field in the 'for patient' window.`n`nClick cancel to exit.","Dispense Date","OC 4096")

            If OKGo != "OK"
                Exit

            WinMatch("Refill for Patient:", 3)
            
            KeyWait "LButton", "D"
            Sleep 500
            }
        Send SUDmap.Get(Integer(A_Index))
        Send "!o"
        WinWaitClose("Refill for Patient:") 
        } until Integer(A_Index) = SUDmap.Count
    }

/* This would be a better method, but WinGetTitle is truncated in Citrix so we can't see Rx# despite its presence in graphical display. DllCall and UIA fail, too. I don't feel like configuring/distributing Tesseract for OCR.
    For k, v IN SUDmap
        {
        WinWaitActive("Refill for Patient:",,15)
        If !WinActive("Refill for Patient:")
            Break
        Else if !WinActive(k)
            Continue
        Else
            {
            If ColorFilter = 2
                {
                Send "#^c" ;Disable color filter.
                Sleep 500
                }
            Loop 15
                {
                Sleep 200
                } until ImageSearch(&xDT, &yDT, 0, 450, 450, 850, "*75 HBITMAP:*" ImageRxDateTime)
            If ColorFilter = 2
                Send "#^c" ;Enable color filter.
            If xDT != "" AND yDT != ""
                MouseClick "left", xDT + 20, yDT + 30
            If xDT = "" OR yDT = ""
                {
                MsgBox("After closing this window, please click into the suspend date/time field to continue this hotkey.","Couldn't Find Suspend Date/Time Field","4096")
                WinWaitActive(SUDmap[WQM["RxNumber"]])
                KeyWait "LButton"
                }
            Send v
            }
        }
*/

; Disable tooltip, notify user, exit, pat on back.
ToolTip
PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\CLICK_Nice.mp3")
MsgBox("Script is complete.","Alt+Shift+F (Failures) Hotkey - SUD","4096 T10")
Exit

; This GUI is used only in the SUD pathway to allow user to enter variables.
GUISUDFail(*)
{
SUDFailGUI := Gui()
SUDFailGUI.Opt("+AlwaysOnTop")
SUDFailGUI.Title := "Failures - Suspend Until Date Adjuster"
SUDFailGUI.AddText("w340", "You selected SUD. This script will compare the assigned Suspend Until Date (SUD) to the future inactivation date. If the difference is less than the limit entered below, the script will mark the Rx for a resuspend.`n`nAfter clicking OK, the script will wait for your first left-click in WQM. Click the row where the script should begin. The script will scroll the list down until it reaches the endpoint defined by your selected sort order.")
SUDFailGUI.AddText("xm y+10","Max SUD Adjustment: ").OnEvent("Click", SUDTT)
SUDFailGUI.AddEdit("r1 vMaxAdj w25 x+2 yp-3", 7)
SUDFailGUI.AddText("x+5 yp+3","List Sort: ").OnEvent("Click", SortByTT)
SUDFailGUI.AddDropDownList("x+5 yp-3 vSortBy w145 BackgroundWhite Choose1",["Last Updated Asc","Failure Detail"])
SUDFailGUI.AddButton("Section w80 x140", "&OK").OnEvent("Click", Submit)
SUDFailGUI.OnEvent("Close", (*) => SUDFailGUI.Destroy())
SUDFailGUI.OnEvent("Escape", (*) => SUDFailGUI.Destroy())
SUDFailGUI.Show "w360"

; Wait for GUI window to close before returning call
WinWaitClose(SUDFailGUI.Hwnd)

; Function to save GUI input
Submit(*)
{
Saved := SUDFailGUI.Submit(1)
If IsInteger(Saved.MaxAdj)
    AND Saved.MaxAdj > 0
    AND Saved.MaxAdj < 31
    {
    MaxDays := Saved.MaxAdj
    Sorted := Saved.SortBy
    }
Else
    Return MaxDays := 0
}

; Define tooltip actions for various events
SortByTT(*)
{
; If tooltip is displayed, disable
If WinExist("ahk_class tooltips_class32")
    {
    ToolTip
    }
Else
    {
    MouseGetPos &xPosMouse, &yPosMouse
    ToolTip "
        (
        In general, Sort by Last Updated Asc (oldest on top) is recommended.
        With this option, the script will stop at the end of your worklist and
        the total runtime should be shorter.

        Sorting by Failure Detail may be preferred if the script has not been run
        recently as this will allow you to screen all 'pure' SUD failures. This
        may result in a longer runtime as previously screened failures will be
        re-evaluated.
        )", xPosMouse, yPosMouse - 100
    SetTimer () => ToolTip(), -10000
    }
}

SUDTT(*)
{
; If tooltip is displayed, disable
If WinExist("ahk_class tooltips_class32")
    {
    ToolTip
    }
Else
    {
    MouseGetPos &xPosMouse, &yPosMouse
    ToolTip "
        (
        This is the maximum difference between current Suspend Until Date
        and projected stop date/expiration date (whichever is earlier).

        Example: SUD set for T+15, expiration is T+10 (no stop date)
        If Max SUD Adjustment >= 5, the script will mark the Rx for resuspend
        If Max SUD Adjustemnt < 5, the script will skip the Rx
        )", xPosMouse, yPosMouse - 100
    SetTimer () => ToolTip(), -10000
    }
}
}
}

FailuresBatchLocalMail(){
GUILocalMailResuspend(&Sta, &MinDate, &MaxDate)

; If variables have not updated, exit script
If !IsSet(MinDate) OR !IsSet(MaxDate) OR !IsSet(Sta)
    Exit

; Display tool tip
ToolTip "Running query...", 875, 32

; Get a list of RxNumbers that are suspended for local mail but appear CMOP eligible
LMRS := LSVQuery(Format("
    (
    SELECT STRING_AGG(RxNumber, ', ')
    FROM LSV.D05_VISN20INFORM.RxMill_OP_Dispense_Open OD
    WHERE OD.Sta3n = '{1}'
        AND OD.RxNumber NOT LIKE '%-[2-4]%'
        AND OD.OrderStatus = 'ORDERED'
        AND OD.RoutingOption LIKE 'LOCAL%'
        AND OD.CreditDispenseHistorySID < 0
        AND OD.LastSuspendStageStatus = 'Suspended/Success'
        AND ISNULL(OD.LocalDispenseStep,0) = 0
        AND ISNULL(OD.RefillDeleted,0) = 0
        AND DispensePriorityDateTime >= CAST(CAST('{2}' AS DATE) AS DATETIME2(0))
        AND DispensePriorityDateTime < CAST(CAST(DATEADD(DD,1,'{3}') AS DATE) AS DATETIME2(0))
        AND NOT EXISTS (SELECT 1
                        FROM LSV.D05_VISN20INFORM.RxMill_OP_Validation_Failure CVF
                        WHERE OD.RebillDispenseHistorySID = CVF.DispenseHistorySID
                            AND CVF.FailureType = 'CANCEL BACK'
                            AND CVF.DispenseHistorySID > 0)
        AND NOT EXISTS (SELECT 1
                        FROM LSV.D05_VISN20INFORM.PatientMill CP
                        WHERE OD.PersonSID = CP.PersonSID
                            AND (CP.ScripTalkFlag = 1
                            OR CP.TestPatientFlag = 1
                            OR CP.DeceasedFlag = 1
                            OR CP.LocalOnlyFlag = 1))   -- Local Mail Only Flag from Cerner_Patient_Address added to LSV view
        AND NOT EXISTS (SELECT 1
                        FROM LSV.D05_VISN20INFORM.Product_Lookup_AHK OPD
                        WHERE OD.ParentItemSID = OPD.ParentItemSID
                            AND OPD.MailOrder != '1')
        AND NOT EXISTS (SELECT 1
                FROM LSV.D05_VISN20INFORM.RxMill_Prescription Rx
                WHERE OD.PersonOrderSID = Rx.PersonOrderSID
                    AND Rx.DAW IN ('1','3','7'))
    GROUP BY Sta6a
    )",Sta3n,MinDate,MaxDate),1)

; If query has no results, disable tooltip, message user, exit
If !LMRS
    {
    ToolTip
    PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\CLICK_Nice.mp3")
    MsgBox("No prescriptions found via query. Exiting script.","Nothing To Do Here","4096 T4")
    Exit
    }

; Display tooltip.
ToolTip "Searching WQM for prescriptions", 875, 32

; Counter tracks if actions are applied so we don't submit zero actions.
RSCount := 0

; Activate and maximize the WQM application as needed.
If !WinActive("Work Queue Monitor - \\\\Remote")
    WinActivate("Work Queue Monitor - \\\\Remote")
If !WinGetMinMax("Work Queue Monitor - \\\\Remote")
    MaxWin("Work Queue Monitor - \\\\Remote")
WinGetPos , , , &hWQM, "Work Queue Monitor - \\\\Remote"
LastRow := hWQM - 122

; Try to find Rx Number column using ImageSearch.
If ColorFilter = 2
    {
    Send "#^c" ;Disable color filter.
    Sleep 500
    }
Loop 3
    {
    ImageSearch(&X, &Y, 0, 0, 1920, 175, "*75 HBITMAP:*" ImageRxNumHeader)
    Sleep 250
    } until X != "" AND Y != ""
If ColorFilter = 2
    Send "#^c" ;Enable color filter.
; If successful, offset to aim click else assume column 2
If X != "" AND Y != ""
    X += 30, Y += 28
If X = "" OR Y = ""
    X := 200, Y := 132

; Click in Rx Number column to enable copy.
MouseClick "left", X, Y
LastRx := "0"
CB := A_Clipboard

Loop
    {
; Increment position if we are not yet at bottom of visible list
    If Y + 18 <= LastRow AND A_Index > 1
        Y += 18

    A_Clipboard := ""
    Send "^c"
    Clipwait(2)
    RxNumber := RegExReplace(A_Clipboard,"\s+")
    If !RxNumber
        {
        If RSCount
            Break
        Else
            {
            ToolTip
            A_Clipboard := CB
            MsgBox("Rx number not found. Exiting script.","Copy Failed","4096 T3")
            Exit
            }
        }
    Else if !RegExMatch(RxNumber,"^\d\d\d\d-\d\d\d\d\d\d\d\d$")
        {
        If RSCount
            Break
        Else
            {
            ToolTip
            A_Clipboard := CB
            MsgBox("Rx number not found. Exiting script.","Format Error","4096 T3")
            Exit
            }
        }

; If RxNumber doesn't change we're at the end of the list, so break this loop
    Else if LastRx = RxNumber
        Break

; If RxNumber exists in result set (Haystack,Needle) then apply the recommended action (resuspend)
    Else if InStr(LMRS,RxNumber)
        {
        Send "{Home}"
        Send "r"
; Enter to prevent allowing down to change action to Release
        Send "{Enter}"
; Sleep here prevents Suspended Order Details Window from triggering
        Sleep 400
; Click back in Rx Number column and increment counter, sleep to allow application to catch up
        MouseClick "left", X, Y
        Sleep 200
        RSCount += 1
        }

    Send "{Down}"
    Sleep 200

    LastRx := RxNumber
    }

A_Clipboard := CB

; No actions were applied, disable tooltip, message user and exit
If !RSCount
    {
    PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\CLICK_Nice.mp3")
    MsgBox("Appears no actions were taken. Exiting script.","Nothing More To Do Here","4096 T4")
    }
Else
    {
    TwoStep := MsgBox(Format("Please review proposed resuspend{1}. The action{1} can be completed manually or by using the Failures => Resuspend to CMOP script.`n`nWould you like to start that script now?",RSCount = 1 ? '' : 's'),"Screening Complete","4096 YN")
    If TwoStep = "Yes"
        {
        WinGetPos , , , &h, "Work Queue Monitor"
        WinActivate("Work Queue Monitor - \\\\Remote")
        ToolTip "Resuspend to CMOP Script is executing", 875, 32

        Send "!s"
        Loop
            {
            WQMResuspend("CMOP","Pharmacist Judgment",1)
            Sleep 3000
            If WinActive("Work Queue Monitor")
                {
                If ColorFilter = 2
                    {
                    Send "#^c" ;Disable color filter.
                    Sleep 500
                    }
                If ImageSearch(&X, &Y, 0, h-30, 50, h, "*75 HBITMAP:*" ImageReady)
                    OR ImageSearch(&X, &Y, 0, h-30, 50, h, "*75 HBITMAP:*" ImageLoading)
                    OR ImageSearch(&X, &Y, 0, h-30, 50, h, "*75 HBITMAP:*" ImageNoRecords)
                    {
                    If ColorFilter = 2
                        Send "#^c" ;Enable color filter.
                    Break
                    }
                If ColorFilter = 2
                    Send "#^c" ;Enable color filter.
                }
            }
        BlockInput "Off"
        BlockInput "MouseMove" "Off"
        ToolTip
        PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\CLICK_Nice.mp3")
        MsgBox("Script is complete","Alt+Shift+F (Failures) Hotkey - Resuspend to CMOP","4096 T10")
        }
    }
ToolTip
Exit

; This GUI is used only in the Local Mail Resuspend pathway to allow user to enter variables to pass to the query
GUILocalMailResuspend(*)
{
LMSusGUI := Gui()
LMSusGUI.Opt("+AlwaysOnTop")
LMSusGUI.Title := "Alt+Shift+F - Local Mail to CMOP"
LMSusGUI.AddText("w340", "You selected Local Mail to CMOP. This script uses a query to identify prescriptions suspended for local mail that may be CMOP eligible.`n`nUpon clicking OK, the query will run then the script will scroll the loaded view to find matching Rx Numbers. A resuspend action will be applied to matching prescriptions.")
LMSusGUI.AddText("","Station Number: ")
LMSusGUI.AddEdit("r1 vSta3n w40 x+5 yp-2", Sta3n)
LMSusGUI.AddText("x+10 yp+2 Section","Min Suspend Until Date: ")
LMSusGUI.AddDateTime("r1 vMinDate w80 xs+120 yp-2 Choose" DateAdd(A_Now,7,"D") " Range" A_Now)
LMSusGUI.AddText("xs y+4","Max Suspend Until Date: ")
LMSusGUI.AddDateTime("r1 vMaxDate w80 xs+120 yp-2 Choose" DateAdd(A_Now,30,"D") " Range" A_Now "-" DateAdd(A_Now,30,"D"))
LMSusGUI.AddButton("Section w80 x140", "&OK").OnEvent("Click", Submit)
LMSusGUI.OnEvent("Close", (*) => LMSusGUI.Destroy())
LMSusGUI.OnEvent("Escape", (*) => LMSusGUI.Destroy())
LMSusGUI.Show "w360"

; Wait for GUI window to close before returning call
WinWaitClose(LMSusGUI.Hwnd)

; Function to save GUI input
Submit(*)
{
Saved := LMSusGUI.Submit(1)
Sta := Saved.Sta3n
MinDate := FormatTime(Saved.MinDate,"ShortDate")
MaxDate := FormatTime(Saved.MaxDate,"ShortDate")
Return
}
}
}

FailuresBatchWindow(){
WinActivate("\QPharmNet: Work Queue Monitor - \\Remote\E")
WinGetPos , , , &h, "Work Queue Monitor"

; Look for Routing Option Reason column in WQM.
If ColorFilter = 2
    {
    Send "#^c" ;Disable color filter.
    Sleep 500
    }
If !ImageSearch(&XROR, &YROR, 0, 115, 1920, 143, "*75 HBITMAP:*" ImageRoutingOptionReason)
    {
    If ColorFilter = 2
        Send "#^c" ;Enable color filter.
    PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\sad trombone.mp3")
    MsgBox("Unable to find Routing Option Reason column. Configure Views to make visible if not already, then try again. Exiting script.","Resuspend Window","4096 T5")
    WinWaitClose("Resuspend Window")
    PlaySound()
    Exit
    }
If ColorFilter = 2
    Send "#^c" ;Enable color filter.

GUIWindowResuspend(&RowCount,&Internal)

If !IsSet(RowCount) OR RowCount = 0
    Exit

; Update click target based on upper left corner of snip and load dummy variable.
XROR += 50, YROR += 30, LastRx := "0"

; Pre-processing checks.
Loop RowCount
    {
    If ColorFilter = 2
        {
        Send "#^c" ;Disable color filter.
        Sleep 500
        }
    While !ImageSearch(&X, &Y, 0, h-30, 50, h, "*75 HBITMAP:*" ImageReady)
        {
        If ImageSearch(&X, &Y, 0, h-30, 50, h, "*75 HBITMAP:*" ImageNoRecords)
            {
            Tooltip
            PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\CLICK_Nice.mp3")
            MsgBox("No records found in WQM. Exiting script.","Nothing To Do Here","4096 T3")
            Exit
            }
        Sleep 250
        }
    If ColorFilter = 2
        Send "#^c" ;Enable color filter.

; Display tool tip with counter.
    ToolTip "Please wait; script is working row " A_Index " of " RowCount ".", 875, 32
    WQM := WQMRow(YROR)

    If !WQM
        {
        YROR += 18
        Continue
        }

    Else if !IsObject(WQM) OR WQM["RxNumber"] = ""
        {
        YROR += 18
        Tooltip "Could not copy RxNumber. Please try again.",,,2
        SetTimer () => ToolTip(,,,2), -3000
        Continue
        }

    Else if WQM["Action"] = "Action"
        {
        ToolTip
        PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\CLICK_Nice.mp3")
        MsgBox("Appears we are at the end of list. Exiting script.","All She Wrote","4096 T3")
        Exit
        }

    Else if WQM["RxNumber"] = LastRx
        {
        A_Index -= 1
        ToolTip "Please wait, script is working row " A_Index " of " RowCount, 875, 32
        }

; If internal refill and user has opted to not resuspend internal, or not suspended for window, skip!
; If not routed for window, skip!
; If not in Suspended status, skip (e.g. release ready)
    If (WQM["RequestMethod"] = "Internal" AND Internal = 0)
        OR WQM["RoutingOption"] != "Window"
        OR WQM["Stage"] != "Suspended"
        {
; Increment Y value for mouseclick.
; Update RxNumber check value.
        YROR += 18
        LastRx := WQM["RxNumber"]
        Continue
        }

; Begin processing.
; Open Routing Option Reason window to check mail status.
    MouseClick "left", XROR, YROR
    WinMatch("Routing Option Validation Results",5)
; If patient is NOT "do not mail", proceed:
    If ColorFilter = 2
        {
        FilterSwitch := 0
        Send "#^c" ;Disable color filter.
        Sleep 500
        }

    Loop 3
        {
        Sleep 250
        ImageSearch(&xd, &yd, 30, 60, 115, 90, "*75 HBITMAP:*" ImageDoNotMail)
        } until xd != "" AND yd != ""

    If xd = "" OR yd = ""
        {
        If ColorFilter = 2
            {
            FilterSwitch := 1
            Send "#^c" ;Enable color filter.
            }
; Close window.
        Send "{Enter}"
        WinMatch("PharmNet: Work Queue Monitor",5)
; Move to Action column, r for Resuspend, !s to submit.
        Send "{Home}"
        Send "r"
        Send "!s"
        WinWaitNotActive("PharmNet: Work Queue Monitor")
; e.g. Release ready, discharged LPE.
; Close window, increment yPos, restart loop.
        If WinActive("Unable to")
            OR WinActive("Profile Actions")
            {
            YROR += 18
            LastRx := WQM["RxNumber"]
            Send "{Enter}"
            WinMatch("PharmNet: Work Queue Monitor",5)
            Continue
            }
        Else if WinActive("Order Discontinued")
            {
            Send "!o" ;OK
            WinMatch("PharmNet: Work Queue Monitor",5)
            MouseClick "left", XROR - 100, YROR
            Send "{Home}"
            Send "d" ;Delete (equivalent to cancel fill in MMR).
            Send "!s" ;Submit
            LastRx := WQM["RxNumber"]
            Continue
            }
        Else if WQM["RequestMethod"] = "Internal"
            {
            WinMatch("Bill-Only",15)
; Navigate to and enter reason.
            Send "{Tab 2}"
            Send "Pharmacist Judgment"
            Send "!o"
; Wait for window to close.
            WinWaitNotActive("Bill-Only")
            Sleep 1000
            }
; Health plan unavailable for this encounter.
        If WinActive("Warning")
            {
            Send "!o"
            WinWaitNotActive("Warning")
            }
; Copay Changed alert.
        If WinActive("Copay Changed")
            {
            Send "!o"
            WinWaitNotActive("Copay Changed")
            }
; Order completed alert - this row will need to be skipped, so update position.
        If WinActive("Order Completed")
            {
            Send "!o"
            WinWaitNotActive("Order Completed")
            YROR += 18
            LastRx := WQM["RxNumber"]
            Continue
            }
        WinMatch("for Patient",15)

; If patient or drug are not CMOP eligible, the routing option update will error.
; If it does, try to update to local mail.
; If that also errors, message user (pauses script) then resume.
        Try RxChangeRouting("CMOP")
        Catch
            {
            Try RxChangeRouting("Local")
            Catch error as err
                {
                LogError(err)
                RouteRx := MsgBox("Unable to confirm successful update to routing option. Please resolve manually, then click OK here.`nCancel to skip.","Where we going?","OC 4096")
                If RouteRx != "OK"
                    {
                    YROR += 18
                    LastRx := WQM["RxNumber"]
                    Send "!c"
                    WinMatch("PharmNet: Work Queue Monitor",5)
                    Continue
                    }
                }
            }
        WinMatch("for Patient",5)
        RxCheckHP()
        Send "!o"
        WinWaitNotActive("for Patient")
; e.g. The stop date must occur before system-calculated stop date.
; This is the same alert that we would get for PRN post 3B but with different text.
; Appears we can copy the text from the window using ^c.
        If WinActive("Validation Error")
            {
            RxValidationError(&YROR)
            If WinWaitActive("for Patient", ,4)
                Send "!c"
            }
        WinMatch("PharmNet: Work Queue Monitor",5)
        }
; If patient is do not mail, update yPos and RxNumber, restart loop.
    Else
        {
        YROR += 18
        LastRx := WQM["RxNumber"]
; Close window.
        Send "{Enter}"
        WinMatch("PharmNet: Work Queue Monitor",15)
        }
    If ColorFilter = 2 AND !FilterSwitch
        Send "#^c" ;Enable color filter.
    Sleep 750
    }
ToolTip
PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\CLICK_Nice.mp3")
MsgBox("Script is complete","Alt+Shift+F (Failures) Hotkey - Window Resuspend","4096 T3")

; This GUI is used only in the Window Resuspend pathway to allow user to enter value and optionally include internal refills.
GUIWindowResuspend(*)
{
WinSusGUI := Gui()
WinSusGUI.Opt("+AlwaysOnTop")
WinSusGUI.Title := "Alt+Shift+F - Window Resuspend"
WinSusGUI.AddText("", "You selected Window. Upon clicking OK, the script will scroll the`nloaded view and attempt to resuspend pending window refills with`na request method other than 'Internal' (MMR).`n`nHow many rows would you like to attempt?")
WinSusGUI.AddEdit("r1 vRowCount w24 x230 y54")
WinSusGUI.AddText("Section x10","To include INTERNAL refills, check this box:")
WinSusGUI.AddCheckBox("vMMRRf x+14")
WinSusGUI.AddButton("Section w80 x130 Default", "&OK").OnEvent("Click", Submit)
WinSusGUI.OnEvent("Close", (*) => WinSusGUI.Destroy())
WinSusGUI.OnEvent("Escape", (*) => WinSusGUI.Destroy())
WinSusGUI.Show "w350"

; Wait for GUI window to close before returning call.
WinWaitClose(WinSusGUI.Hwnd)

; Function to save GUI input.
Submit(*)
{
Saved := WinSusGUI.Submit(1)
; If RowCount value is non-integer, leave it as zero so script will exit.
If IsInteger(Saved.RowCount) AND Saved.RowCount > 0
    {
    Internal := Saved.MMRRf
    RowCount := Saved.RowCount
    }
Else
    Return RowCount := 0
Return
}
}
}
}
#HotIf

/****************************************************** Ctrl+Left Mouse Button  (Claim Reject) *****************************************************
; This hotkey is used to resolve individual rejected claims in Claims Monitor.
****************************************************************************************************************************************************/

#HotIf WinActive("PharmNet: Claims Monitor - \\\\Remote")
^LButton::
{
; Wait for ctrl to release then capture mouse pos and pass to CMRow.
KeyWait "ctrl", "T2"
MouseGetPos(&xPosMouse,&yPosMouse)
CM := CMRow(yPosMouse)

If CM["Transaction"] = "Reverse" AND CM["HxFill"]
    {
    Affirm := MsgBox("This is a historical bill. If you would like the script to reverse it, click OK here.`n`nPreview comment: 'Historical bill from " CM["ServiceDate"] ".'`n`nClick cancel to exit.","Historical Bill Reversal","OC 4096")

    If Affirm != "OK"
        Exit

    CMReverse(CM, yPosMouse)
    PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\CLICK_Nice.mp3")
    MsgBox("Historical bill has been reversed. Click Submit to accept, or Cancel to undo. Then refresh the view.","Reversal Complete","4096 T5")
    }
Else if CM["Transaction"] = "Reverse" AND !CM["HxFill"]
    MsgBox("This reversal is not for a historical bill. Proceed manually.","Manual Intervention Required","4096 T5")
Else
    {
; Update user about what you're going to do.
    If CM["DummyHPFlag"] = 0 OR CheckHPOn = 2
        ToolTip 'Initiating unclaim action...', xPosMouse, yPosMouse, 1
    Else
        ToolTip 'Initiating removal of DoD Health Plan...', xPosMouse, yPosMouse, 1

    SetTimer () => ToolTip(), -3000

; Do that and submit if we unclaimed.
    CMReject(CM, yPosMouse)
    If CM["DummyHPFlag"] = 0 OR CheckHPOn = 2
        Send "!m"

    PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\CLICK_Nice.mp3")
    }
}
#HotIf

/****************************************************** Ctrl+Left Mouse Button  (Convert Rx) *******************************************************
; This hotkey is used to complete Rx conversions.
; Conversions are defined using a map, similar to LDFMap.
****************************************************************************************************************************************************/

#HotIf WinActive("- PharmNet: Ret") AND MMRCtrlLButton = 1
^LButton::{
KeyWait "Control", "T2"

Rx := MMRRxRow()

If !IsObject(Rx)
    Exit

ConvertRx(Rx)
}
#HotIf

#HotIf WinActive("- PharmNet: Ret")

/****************************************************** Ctrl+Right Mouse Button (Hold PADR Intervention) ******************************************
; This hotkey is used to ready a hold intervention for PADR from MMR.
; X-Rx (or presumed) orders will be prompted for subsequent placement of PADR consult in PowerChart.
; Ultimately hoping to make it easier for frontline pharmacists to initiate consluts.
***************************************************************************************************************************************************/

^RButton::{
KeyWait "Control", "T2"
WinGetPos(&X,&Y,,,"A")
If WinExist("Pharmacy Eligibility Overview")
    WinClose("Pharmacy Eligibility Overview")
PADRx := MMRRxRow() ;Get Source Type and Order Sentence.
PatientName := RegExReplace(WinGetTitle("A"),"i) - Pharm.*\Q\\Remote\E") ;Needed to identify windows.
Tooltip "Preparing Hold Rx Documentation", 881, 32
IholdPADR(PADRx, Y, PatientName)
PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\CLICK_Nice.mp3")
Tooltip "Hold Rx Documentation prepared for " PatientName ": Please review and sign.", 881, 32
SetTimer () => ToolTip(), -3000
;C-Rx orders should already have a consult entered by provider in PowerChart.
If !InStr(PADRx["SourceType"], "C") ;InStr used here in case of MMRRxRow fail.
    {
    WinWaitActive(PatientName " - PharmNet")
    CCPC := MsgBox("If this is a medication from CitC needing NF/PADR adjudication, a consult must be entered into PowerChart.`n`nClick OK to continue, or Cancel to exit.","External NF/PADR consult?","OC 4096")
    If CCPC != "OK"
        Exit
    If !WinActive(PatientName " - PharmNet")
        WinActivate(PatientName " - PharmNet")
    Sleep 250
    Send "^e" ;Launch PowerChart!
    KeyWait "Control", "T2"
    ;Try
    ;    {
    ;    WinActivate(PatientName " - PharmNet")
    ;    Send "^p" ;Close LPE in MMR to avoid apparent conflict in PowerChart.
    ;    Tooltip "Don't worry, I'll open it again later... fEHR at its finest.",,,2
    ;    SetTimer () => ToolTip(,,,2), -3000
    ;    }
    Tooltip "Opening PowerChart", 881, 32
    If WinExist("Pharmacy Eligibility Overview")
        WinClose("Pharmacy Eligibility Overview")
    LastName := StrSplit(PatientName,",",,2) ;Suffixes aren't consistent.
    Loop 10
        {
        Sleep 500 ;Wait <= 5s for Veteran's profile to load in PowerChart.
        } until WinExist(LastName[1] ".*" LastName[2] " -.*Opened by.*\Q\\Remote\E") OR WinExist("Discern:")
    While WinExist("Discern:")
        {
        If !WinActive("Discern:")
            WinActivate("Discern:")
        If WinActive("Discern:")
            {
            Send "{Enter}"
            WinWaitClose("Discern:")
            }
        WinWait("Discern:",,2)
        }
    Loop 10
        {
        Sleep 500 ;Wait <= 5s for Veteran's profile to load in PowerChart.
        } until WinExist(LastName[1] ".*" LastName[2] " -.*Opened by.*\Q\\Remote\E") OR WinExist("i)Assign a Relationship")
    If WinExist("i)Assign a Relationship") AND !WinActive("i)Assign a Relationship")
        WinActivate("i)Assign a Relationship")
    If WinActive("i)Assign a Relationship")
        {
        WinWaitClose("i)Assign a Relationship")
        WinWait(LastName[1] ".*" LastName[2] " -.*Opened by.*\Q\\Remote\E")
        }
    If WinExist(LastName[1] ".*" LastName[2] " -.*Opened by.*\Q\\Remote\E") AND !WinActive(LastName[1] ".*" LastName[2] " -.*Opened by.*\Q\\Remote\E")
        {
        Try
            WinActivate(LastName[1] ".*" LastName[2] " -.*Opened by.*\Q\\Remote\E") ;Can open in background if other PowerChart tabs exist.
        Catch
            {
            APC := MsgBox("Waiting for Veteran's profile to load in PowerChart. Please click OK to proceed (window is now loaded) or Cancel to exit. If you exit, the next hotkey can be triggered with Ctrl+Lclick from the Veteran's profile in PowerChart.","Window Not Found","OC 4096")
            If APC != "OK"
                {
                Tooltip
                Exit
                }
            If !WinActive(LastName[1] ".*" LastName[2] " -.*Opened by.*\Q\\Remote\E")
                {
                Try
                    WinActivate(LastName[1] ".*" LastName[2] " -.*Opened by.*\Q\\Remote\E")
                Catch ;Not callin' the end user a liar, but I ain't necessarily callin' 'em a truther!
                    {
                    Tooltip
                    PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\sad trombone.mp3")
                    MsgBox("Script failed to locate the Veteran's profile on PowerChart. Please continue manually or trigger the next hotkey with Ctrl+Lclick in the Veteran's profile on PowerChart.`n`nNow exiting script.","Can't Find PowerChart Window","4096 T5")
                    WinWaitClose("Can't Find PowerChart Window")
                    PlaySound()
                    Exit
                    }
                }
            }
        }
    Else if !WinExist(LastName[1] ".*" LastName[2] " -.*Opened by.*\Q\\Remote\E") ;When Cerner's bein' slow fr.
        {
        VPC := MsgBox("Waiting for Veteran's profile to load in PowerChart. Please click OK to proceed (window is now loaded) or Cancel to exit. If you exit, the next hotkey can be triggered with Ctrl+Lclick from the Veteran's profile in PowerChart.","Window Not Found","OC 4096")
        If VPC != "OK"
            {
            Tooltip
            Exit
            }
        If !WinActive(LastName[1] ".*" LastName[2] " -.*Opened by.*\Q\\Remote\E")
            {
            Try
                WinActivate(LastName[1] ".*" LastName[2] " -.*Opened by.*\Q\\Remote\E")
            Catch ;Not callin' the end user a liar, but I ain't necessarily callin' 'em a truther!
                {
                Tooltip
                PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\sad trombone.mp3")
                MsgBox("Script failed to locate the Veteran's profile on PowerChart. Please continue manually or trigger the next hotkey with Ctrl+Lclick in the Veteran's profile on PowerChart.`n`nNow exiting script.","Can't Find PowerChart Window","4096 T5")
                WinWaitClose("Can't Find PowerChart Window")
                PlaySound()
                Exit
                }
            }
        }
    Sleep 2500
    Tooltip "Adding BVE", 881, 32
    PCAddBVE(LastName[1] ".*" LastName[2] " -.*Opened by.*\Q\\Remote\E")
    WinWait("i)Assign a Relationship",,2)
    If WinExist("i)Assign a Relationship") AND !WinActive("i)Assign a Relationship")
        WinActivate("i)Assign a Relationship")
    If WinActive("i)Assign a Relationship")
        WinWaitActive(LastName[1] ".*" LastName[2] " -.*Opened by.*\Q\\Remote\E")
    Tooltip "Ordering Consult", 881, 32
    PCNewOrder("Outside Provider", PADRx["Prescriber"], LastName[1] ".*" LastName[2] " -.*Opened by.*\Q\\Remote\E", LastName[1] ".*" LastName[2])
    Tooltip "Preparing Consult", 881, 32
    InitiateNFPA(PADRx["OrderSentence"],PADRx["RxNumber"],LastName[1] ".*" LastName[2] " -.*Opened by.*\Q\\Remote\E")
    PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\CLICK_Nice.mp3")
    }
}
#HotIf

/****************************************************** Ctrl+Left Mouse Button - WQM  **************************************************************
; This hotkey is used to work the things in various views of Work Queue Monitor
; Multiple pathways exist:
; - Surescripts cancel requests can be called independently using ctrl+left-click or via GUI
; - Phone notify + Surescripts IN PROCESS rows can be deleted as they are complete
; - Update qty/days supply for items with CMOP/PBM restriction
****************************************************************************************************************************************************/

#HotIf WinActive("PharmNet: Work Queue Monitor")
^LButton::
{
KeyWait "Control", "T2"

MouseGetPos , &yPosMouse

WQM := WQMRow(yPosMouse)

If WQM["TransactionType"] = "Change" AND InStr(WQM["AuthorizationStatus"], "NOTIFIED")
    {
    RegExMatch(WQM["LastUpdated"], "(\d\d)/(\d\d)/(\d\d\d\d) (\d\d)\:(\d\d)", &Chronon)
    UpdatedLast := String(Format("{:04}{:02}{:02}{:02}{:02}{:02}", Chronon[3], Chronon[1], Chronon[2], Chronon[4], Chronon[5], 0))
    If DateDiff(A_Now,UpdatedLast,"days") ;Daily notifications, max.
        {
        Send "{Home}" ;Action column.
        SendText "n" ;Notify
        Send "{Enter}" ;Confirm
        Send "!s" ;Submit
        PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\CLICK_Nice.mp3")
        Tooltip "Reminder sent to provider.",,,2
        SetTimer () => ToolTip(,,,2), -3000
        Exit
        }
    }

Else if WQM["TransactionType"] = "Change"
    {
    PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\refuse.mp3")
    Tooltip "Hotkeys not yet developed for RxChange requests, but maybe someday!",,,2
    SetTimer () => ToolTip(,,,2), -3000
    Exit
    }

Else if InStr(WQM["FailureDetail"], "Cancel request exists for order") OR InStr(WQM["FailureDetail"], "Patient is deceased")
    {
    Send "{Home}" ;Action column.
    SendText "d" ;Delete
    Send "{Enter}" ;Confirm
    Send "!s" ;Submit
    If WQM["RequestMethod"] = "Internal" ;MMR proxy will trigger to cancel fill.
        {
        WinMatch("(Cancel-Fill|Bill-Only) " WQM["RxNumber"], 9)
        Send "{tab 2}"
        Send "Prescriber Modified Order"
        Send "!o"
        WinWaitNotActive("Cancel-Fill " WQM["RxNumber"])
        }
    RegExMatch(WQM["LastUpdated"], "(\d\d)/(\d\d)/(\d\d\d\d) (\d\d)\:(\d\d)", &Chronon)
    UpdatedLast := String(Format("{:04}{:02}{:02}{:02}{:02}{:02}", Chronon[3], Chronon[1], Chronon[2], Chronon[4], Chronon[5], 0))
    If DateDiff(A_Now,UpdatedLast,"days") > 13
        {
        WinMatch("Cancel Fill Warning", 9) ;Will fire if the Rx has been in the queue for two weeks or more.
        SchedCheck := Integer(SubStr(WQM["RxNumber"], 6, 1)) ;This digit will be <6 for controlled substances.
        If !InStr(WQM["RoutingOption"], "CMOP") AND (SchedCheck < 5 OR RegExMatch(WQM["LegalStatus"], "\d+")) ;Cancelling old fills for controlled substances may disturb local inventory.
            {
            PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\bell.wav")
            ContCheck := MsgBox("This Rx is for a controlled substance, and was filled locally a while ago.`n`nCancelling this fill will increment inventory, so please ensure counts remain accurate.`n`nClick 'OK' to proceed, or 'Cancel' to exit.","Warning: Inventory Adjustment","OC 4096")
            If ContCheck != "OK"
                {
                WinMatch("Cancel Fill Warning",3)
                Send "!c" ;cancel warning
                If WQM["RequestMethod"] = "Internal"
                    {
                    WinMatch("Cancel-Fill " WQM["RxNumber"], 5)
                    Send "!c" ;cancel fill cancel
                    }
                Exit
                }
            WinMatch("Cancel Fill Warning",3)
            }
        Send "!o"
        }
    PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\CLICK_Nice.mp3")
    Tooltip "Fill cancelled. Next, execute the 'Surescripts Cancel' hotkey (Alt+Shift+E) from 'eRx Requests' view.",,,2 ;eRx Requests button from startup GUI.
    SetTimer () => ToolTip(,,,2), -5000
    Exit
    }
Else if WQM["FailureDetail"] = "012:SITE REQUESTED CANCEL"
    {
    If !WinExist("PharmNet: Retail")
        {
        MsgBox("This script requires both Work Queue Monitor (WQM) and Medication Manager Retail (MMR). Unable to find MMR. Please launch application and retry hotkey.","Application Check","4096 T5")
        Exit
        }
    
    ToolTip 'Initiating Rx Cancel'
    SetTimer () => ToolTip(), -3000

    MMRPatientLookup(WQM["RxNumber"],"Rx")
    ;WinWaitNotActive("^PharmNet: Retail", ,5)
    If !MMRFindRx(WQM["RxNumber"])
        {
        MsgBox(Format("Search complete. Rx#: {1} not found on active list.", WQM["RxNumber"]),"Rx Search Complete","T3")
        Exit
        }

; Delete in WQM doesn't work for CMOP cancel backs, so we have to cancel in MMR.
    Send "{Home}"
    SendText "c"
    Send "{Enter}"
    Send "!a"
    KeyWait "Alt", "T2"
    WQMCancelFill(1,'CMOP Cancellation')
    Send "!s"
    Exit
    }

ListLines false
Try Loop LDFMap.Count
    {
    If !LDFMap[A_Index].Has("Item Description") OR !LDFMap[A_Index].Has("Notification Type") OR !LDFMap[A_Index].Has("Status")
        Continue
    If InStr(WQM["OrderSentence"],LDFMap[A_Index]["Item Description"]) AND LDFMap[A_Index]["Notification Type"] = "Hotkey Convertible" AND LDFMap[A_Index]["Status"] = "On"
        {
        ConversionDrug := LDFMap[A_Index]["Item Description"]
        ConversionDetail := LDFMap[A_Index]["Message"]
        Break
        }
    }
ListLines true

; Check if we are in eRx Requests view and are on a row we can delete.
If (WQM["RequestMethod"] = "Phone" AND WQM["TransactionType"] = "Renewal") OR (WQM["TransactionType"] = "Cancel" AND WQM["AuthorizationStatus"] = "In Process")
    {
    ToolTip 'Deleting completed row from eRx Requests view'
    SetTimer () => ToolTip(), -3000
    Send "{Home}" ;Navigate to action cell of selected row.
    SendText "d" ;Delete
    Send "{Enter}"
    Send "!s"
    }

; or in eRx Requests view (AuthStatus is only populated by WQMRow for this view)
Else if WQM.Has("AuthorizationStatus")
    eRxCancel(WQM,yPosMouse)

Else if InStr(WQM["FailureDetail"], "product has been inactivated") AND WQM["Action"] != "Action" ;AND WQM["RequestMethod"] != "Internal"
    {
    ToolTip "Reselecting active product."
    SetTimer () => ToolTip(), -3000
    WQMFail(WQM)
    }

; Next, check if we are looking at a convertible Rx in a Suspended Orders view.
Else if IsSet(ConversionDrug) AND IsSet(ConversionDetail) AND WQM["Stage"] = "Suspended" AND (WQM["Status"] = "Success" OR WQM["FailureDetail"] = "Product is not available at CMOP" OR WQM["FailureDetail"] = "The product has been inactivated") ;B13 appears to have removed WQM's ability to handle these itself.
    {
    If !WinExist("PharmNet: Retail")
        {
        MsgBox("This script requires both Work Queue Monitor (WQM) and Medication Manager Retail (MMR). Unable to find MMR. Please launch application and retry hotkey.","Application Check","4096 T5")
        Exit
        }

    ToolTip 'Initiating Rx conversion for ' ConversionDrug
    SetTimer () => ToolTip(), -3000

; Remove pending refill - this is a 'pre-transmission' refill so expectation is refill has not been decremented from display in Amb Profile
; Internal refills will have a dispense and cancel fill to complete.

    SendText "d"
    Send "!s"
    KeyWait "Alt", "T2"
    If WQM["RequestMethod"] = "Internal"
        WQMCancelFill(1,'Formulary Substitution')

; Find rx on Amb profile, capture details as map, run conversion
    MMRPatientLookup(WQM["RxNumber"],"Rx")
    WinWaitNotActive("^PharmNet: Retail", ,5)
    If !MMRFindRx(WQM["RxNumber"])
        {
        MsgBox(Format("Search complete. Rx#: {1} not found on active list.", WQM["RxNumber"]),"Rx Search Complete","T3")
        Exit
        }

    Rx := MMRRxRow()
    Rx.Set("DPD",WQM["SUD"])

    If !IsObject(Rx)
        Exit

    ConvertRx(Rx,ConversionDrug,ConversionDetail)
    }

; Or an rx with a need for days supply adjustment
Else if IsSet(ConversionDrug) AND IsSet(ConversionDetail) AND WQM["FailureDetail"] = "003:PBM/CMOP DISPENSING LIMIT EXCEEDED"
    {
    If !WinExist("PharmNet: Retail")
        {
        MsgBox("This script requires both Work Queue Monitor (WQM) and Medication Manager Retail (MMR). Unable to find MMR. Please launch application and restart script.","Application Check","4096 T5")
        Exit
        }

    ToolTip 'Initiating Rx replacement'
    SetTimer () => ToolTip(), -3000

    MMRPatientLookup(WQM["RxNumber"],"Rx")
    WinWaitNotActive("^PharmNet: Retail", ,5)
    If !MMRFindRx(WQM["RxNumber"],2)
        {
        MsgBox(Format("Search complete. Rx#: {1} not found on active list.", WQM["RxNumber"]),"Rx Search Complete","T3")
        Exit
        }
    Rx := MMRRxRow(WQM["RxNumber"])
    Rx.Set("DPD",WQM["SUD"])

; Delete in WQM doesn't work for CMOP cancel backs, so we have to cancel in MMR.
    SendText "c"
    Send "{Enter}"
    Send "!a"
    KeyWait "Alt", "T2"
    WQMCancelFill(1,'Wrong Quantity')
    Send "!s"

    If ColorFilter = 2
        {
        Send "#^c" ;Disable color filter.
        Sleep 1000
        }
    Loop 20
        {
        ImageSearch(&X, &Y, 600, 200, 1900, 240, "*75 HBITMAP:*" ImageOrderSentence)
        Sleep 250
        } until X != "" AND Y != ""
    If ColorFilter = 2
        Send "#^c" ;Enable color filter.

    If X = "" OR Y = ""
        {
        APCheck := MsgBox("Unable to confirm cancel fill action is complete and MMR is ready. Click OK if script can proceed or Cancel to exit.","4096 OC")
        If APCheck != "OK"
            Exit
        WinMatch("- PharmNet: Retail Med",5)
        }

; After cancel is successful, update RxMap to reflect credited refill.
    Rx["RR"] += 1

    ConvertRx(Rx)
    }

; Check if we are in a Suspended Orders Failure view and failure qualifies for Resuspend.

/*
001:TEMP OUT OF STOCK
001:TEMP OUT OF STOCK-OPEN MARKET
007:MANUFACTURER'S BACKORDER
013:TEMP OUT OF STOCK-PRIME VENDOR
015:REFRIG ITEM TO PO BOX FILL LOCALLY
016:FLAMMABLE ITEM SHIPPING RESTRICTION
020:DATABASE ERROR-PLEASE RESUBMIT
020:- DATABASE ERROR-PLEASE RESUBMIT
039:PLEASE RESUBMIT FOR REROUTING
Contact Support: Unable to transmit to CMOP
Patient mailing preference is not eligible for CMOP
Product is not available at CMOP
Health plan either ineffective or card holder identity changed
Ineligible health plan found due to benefit type allocation
*/

Else if WQM["Action"] != "Action" AND !InStr(WQM["FailureDetail"], ",") AND (
    InStr(WQM["FailureDetail"], "OUT OF STOCK")
    OR InStr(WQM["FailureDetail"], "PLEASE RESUBMIT")
    OR InStr(WQM["FailureDetail"], "FILL LOCAL")
    OR InStr(WQM["FailureDetail"], "SHIPPING RESTRICT")
    OR InStr(WQM["FailureDetail"], "BACKORDER")
    OR InStr(WQM["FailureDetail"], "Unable to transmit to CMOP")
    OR RegExMatch(WQM["FailureDetail"], "not .* CMOP")
    OR Trim(WQM["FailureDetail"]) = "Health plan either ineffective or card holder identity changed"
    OR Trim(WQM["FailureDetail"]) = "Ineligible health plan found due to benefit type allocation"
)
    {
    ToolTip "Initiating resuspend action for: " WQM["FailureDetail"]
    SetTimer () => ToolTip(), -3000
    WQMFail(WQM)
    }

; Or if it qualifies for Modify
/* Removed this option 06/04/2025 to align with recommendation of avoiding modify
Else if WQM["FailureDetail"] ~= Modify
    AND WQM["Action"] != "Action"
    {
    ToolTip "Initiating modify action for: " WQM["FailureDetail"]
    SetTimer () => ToolTip(), -3000
    WQMFail(WQM)
    }
*/

; Or maybe this is a window order and we want to resuspend it to CMOP
Else if WQM["Status"] = "Success" AND WQM["Stage"] = "Suspended" AND WQM["RoutingOption"] = "Window"
    {
    ToolTip "Initiating CMOP resuspend for Window Suspended Rx."
    SetTimer () => ToolTip(), -3000

    If ColorFilter = 2
        {
        Send "#^c" ;Disable color filter.
        Sleep 500
        }
    If !ImageSearch(&XROR, &YROR, 0, 115, 1920, 143, "*75 HBITMAP:*" ImageRoutingOptionReason)
        {
        If ColorFilter = 2
            Send "#^c" ;Enable color filter.
        ToolTip "Unable to find Routing Option Reason column. Please try again."
        SetTimer () => ToolTip(), -3000
        Exit
        }
    If ColorFilter = 2
        Send "#^c" ;Enable color filter.

    ; Click to open Routing Option Reason window.
    MouseClick "left", XROR + 50, yPosMouse
    WinMatch("Routing Option Validation Results",5)
; If patient is NOT "do not mail", proceed:
    If ColorFilter = 2
        {
        FilterSwitch := 0
        Send "#^c" ;Disable color filter.
        Sleep 500
        }

    Loop 3
        {
        Sleep 250
        ImageSearch(&xd, &yd, 30, 60, 115, 90, "*75 HBITMAP:*" ImageDoNotMail)
        } until xd != "" AND yd != ""

    If xd = "" OR yd = ""
        {
        If ColorFilter = 2
            {
            FilterSwitch := 1
            Send "#^c" ;Enable color filter.
            }
; Close window.
        Send "{Enter}"
        WinMatch("PharmNet: Work Queue Monitor",5)
; Move to Action column, r for Resuspend, !s to submit.
        Send "{Home}"
        SendText "r"
        Send "!s"
        WQMResuspend("CMOP","Pharmacist Judgment",2)
        }
    If ColorFilter = 2 AND !FilterSwitch
        Send "#^c" ;Enable color filter.
    Else
        {
        Send "{Enter}"
        ToolTip 'Patient is Do Not Mail, unable to suspend for CMOP.'
        SetTimer () => ToolTip(), -3000
        WinMatch("PharmNet: Work Queue Monitor",5)
        }
    }

; Let the user know you tried.
Else ToolTip "Unable to determine what to do. Please try again."

SetTimer () => ToolTip(), -3000
}
#HotIf

/******************************************* eRx Requests - Batch action hero ****************************************************
; This hotkey is used to work batches of eRx Requests view in Work Queue Monitor.
**********************************************************************************************************************************/

#HotIf WinExist("PharmNet: Work Queue Monitor")
!+e::
eRxRequests(hk)
{
KeyWait "Alt", "T2"

GUIeRxRequest(&eRxPath)

If !IsSet(eRxPath)
    Exit

If !WinActive("PharmNet: Work Queue Monitor")
    WinActivate("PharmNet: Work Queue Monitor")
If !WinGetMinMax("PharmNet: Work Queue Monitor")
    MaxWin("PharmNet: Work Queue Monitor")

If eRxPath = "CancelRx"
    {
    If !WinExist("PharmNet: Retail")
        {
        MsgBox("This script requires both Work Queue Monitor (WQM) and Medication Manager Retail (MMR). Unable to find MMR. Please launch application and restart script.","Application Check","4096 T5")
        Exit
        }
    OKGo := InputBox("To begin the script, click OK here then click on the row where you would like the script to start. The script will start one second after the first left-click in Work Queue Monitor.`n`nHow many rows would you like to attempt?","eRx Cancel Requests")
    If OKGo.Result != "OK" OR IsInteger(OKGo.Value) < 1
        Exit

    WinActivate("PharmNet: Work Queue Monitor")
    WinGetPos , , , &h, "Work Queue Monitor"
    Send "^n" ;eRx Requests view.
    KeyWait "LButton", "D"
    MouseGetPos &xPosMouse, &yPosMouse
    Sleep 1000

    Loop OKGo.Value
        {
        If A_Index = 1
            LastWQMRx := ""
        Else if WQM AND IsObject(WQM)
            LastWQMRx := WQM["RxNumber"]
        ; Click row header to highlight row, then extract Rx Number from row.
        WQM := WQMRow(yPosMouse)
        ToolTip "Processing row " A_Index " of " OKGo.Value ".", 881, 32
        CancelResult := eRxCancel(WQM, yPosMouse, LastWQMRx)
        If !CancelResult
            yPosMouse += 18
        Else if CancelResult = 2
            {
            yPosMouse += 18
            A_Index -= 1
            }
        Else
            {
            If ColorFilter = 2
                {
                Send "#^c" ;Disable color filter.
                Sleep 500
                }
            While !ImageSearch(&X, &Y, 0, h-30, 50, h, "*75 HBITMAP:*" ImageReady)
                Sleep 500 ;Until finished submitting.
            If ColorFilter = 2
                Send "#^c" ;Enable color filter.
            }
        }
    ToolTip
    }

Else if eRxPath = "DeleteRow"
    {
    yPosMouse := 155
; Display tool tip so user knows script is running.
    ToolTip "Please wait, script is executing.", 881, 32
; Capture window height so we can run to the bottom.
    WinActivate("PharmNet: Work Queue Monitor")
    WinGetPos , , , &h, "Work Queue Monitor"
    DeleteCount := 0
    NotifyCount := 0
    Send "^n" ;eRx Requests view.
    If ColorFilter = 2
        {
        Send "#^c" ;Disable color filter.
        Sleep 500
        }
    While !ImageSearch(&X, &Y, 0, h-30, 50, h, "*75 HBITMAP:*" ImageReady)
        Sleep 500 ;Until finished loading.
    If ColorFilter = 2
        Send "#^c" ;Enable color filter.
    Loop
        {
        If !WinActive("PharmNet: Work Queue Monitor")
            WinActivate("PharmNet: Work Queue Monitor")
        If A_Index = 1
            LastWQMRx := "", LastWQMLeftovers := "", LastLastUpdated := ""
        Else if WQM AND IsObject(WQM)
            LastWQMRx := WQM["RxNumber"], LastWQMLeftovers := WQM["Leftovers"], LastLastUpdated := WQM["LastUpdated"]

        Sleep 250
        WQM := WQMRow(yPosMouse)

; If clip failed.
        If (!IsObject(WQM) OR !WQM) AND yPosMouse < h - 120
            {
            yPosMouse += 18
            If WQM AND !IsObject(WQM)
                {
                ToolTip("Unable to parse copied row. Try again later.",,2)
                SetTimer () => ToolTip(,,,2), -3000
                }
            Continue
            }

; or copied value = delete +/- preceding tab indicates row was worked or we are at end of list
; clip unchanged indicates end of list
        If WQM["Action"] ~= "(Action|Delete)" OR (WQM["RxNumber"] = LastWQMRx AND WQM["Leftovers"] = LastWQMLeftovers AND WQM["LastUpdated"] = LastLastUpdated)
            Break

        If WQM["TransactionType"] = "Change" AND InStr(WQM["AuthorizationStatus"], "NOTIFIED")
            {
            RegExMatch(WQM["LastUpdated"], "(\d\d)/(\d\d)/(\d\d\d\d) (\d\d)\:(\d\d)", &Chronon)
            UpdatedLast := String(Format("{:04}{:02}{:02}{:02}{:02}{:02}", Chronon[3], Chronon[1], Chronon[2], Chronon[4], Chronon[5], 0))
            If DateDiff(A_Now,UpdatedLast,"days") ;Daily notifications, max.
                {
                MouseClick "left", 65, yPosMouse
                Sleep 250
                SendText "n"
                Sleep 250
                Send "{Enter}"
                KeyWait "Enter", "T2"
                NotifyCount += 1
                }
            yPosMouse += 18
            }

; We don't want to include Surescripts renewals.
        Else if ((WQM["RequestMethod"] = "Phone" AND WQM["TransactionType"] = "Renewal") OR (WQM["TransactionType"] = "Cancel" AND WQM["AuthorizationStatus"] = "In Process"))
            {
; Click in action cell of selected row.
; Delete for delete.
            MouseClick "left", 65, yPosMouse
            Sleep 250
            SendText "d" ;Sometimes applies the Label action? (Adding delay to see if that helps.)
            Sleep 250
            Send "{Enter}"
            KeyWait "Enter", "T2"
; Increment counter and position.
            DeleteCount += 1
            yPosMouse += 18
            }
        Else yPosMouse += 18
        } until yPosMouse >= h - 120

    If WQM AND IsObject(WQM)
        LastWQMRx := WQM["RxNumber"], LastWQMLeftovers := WQM["Leftovers"], LastLastUpdated := WQM["LastUpdated"]
    
    yPosMouse -= 18
    Send "{down}"
    Sleep 250
    WQM := WQMRow(yPosMouse)

    While ((WQM["RxNumber"] AND WQM["RxNumber"] != LastWQMRx) OR (!WQM["RxNumber"] AND WQM["Leftovers"] != LastWQMLeftovers)) AND WQM["Action"] != "Delete" AND WQM["Action"] != "Notify"
        {
        If !IsObject(WQM) OR !WQM
            {
            If WQM AND !IsObject(WQM)
                {
                ToolTip("Unable to parse copied row. Try again later.",,2)
                SetTimer () => ToolTip(,,,2), -3000
                }
            }
        Else if WQM["TransactionType"] = "Change" AND InStr(WQM["AuthorizationStatus"], "NOTIFIED")
            {
            RegExMatch(WQM["LastUpdated"], "(\d\d)/(\d\d)/(\d\d\d\d) (\d\d)\:(\d\d)", &Chronon)
            UpdatedLast := String(Format("{:04}{:02}{:02}{:02}{:02}{:02}", Chronon[3], Chronon[1], Chronon[2], Chronon[4], Chronon[5], 0))
            If DateDiff(A_Now,UpdatedLast,"days") ;Daily notifications, max.
                {
                MouseClick "left", 65, yPosMouse
                Sleep 250
                SendText "n"
                Sleep 250
                Send "{Enter}"
                KeyWait "Enter", "T2"
                NotifyCount += 1
                }
            }
        Else if (WQM["RequestMethod"] = "Phone" AND WQM["TransactionType"] = "Renewal") OR (WQM["TransactionType"] = "Cancel" AND WQM["AuthorizationStatus"] = "In Process")
            {
            MouseClick "left", 65, yPosMouse
            Sleep 250
            SendText "d" ;Sometimes applies the Label action? (Adding delay to see if that helps.)
            Sleep 250
            Send "{Enter}"
            KeyWait "Enter", "T2"
            DeleteCount += 1
            }
        If WQM AND IsObject(WQM)
            LastWQMRx := WQM["RxNumber"], LastWQMLeftovers := WQM["Leftovers"], LastLastUpdated := WQM["LastUpdated"]
        Send "{down}"
        Sleep 250
        WQM := WQMRow(yPosMouse)
        }

    If DeleteCount OR NotifyCount
        {
        Sleep 250
        Send "!s"
        }

; Turn off tooltip.
    ToolTip
    PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\CLICK_Nice.mp3")
    MsgBox(DeleteCount " deletion(s) and " NotifyCount " notification(s) submitted.","eRx Request Matching","4096 T5")
    }

GUIeRxRequest(*)
{
eRxGUI := Gui()
eRxGUI.Opt("+AlwaysOnTop")
DocLinkText := 'Please see <a href="https://dvagov.sharepoint.com/sites/PharmacyCernerEHRCommunityofPractice/SitePages/Oracle-Cerner-Repository-for-Automating-Pharmacy-VA.aspx?csf=1&web=1&share=EZYFbCtdCr9Bvseh7gKT8tMB4MMSjq_qW1N5PuvnfEAATw&e=o2SoTo&cid=cfb91149-243b-4ac4-b714-d4e64239a0ad#description-2">documentation</a> for more details.'
eRxGUI.SetFont("s9.5")
eRxGUI.Title := "Alt+Shift+E (eRx Requests) Hotkey"
eRxGUI.AddText("Section","Use this option to complete pending Surescripts cancel request(s) from the eRx Requests`nview in Work Queue Monitor. The script will discontinue the prescription if active,`nthen return to Work Queue Monitor to approve the cancel request.")
eRxGUI.AddText("y+8","NOTE: Multiple rows can be processed from here. Individual rows can be worked by`nusing Ctrl+Left-Click from the eRx Requests view in Work Queue Monitor.")
eRxGUI.AddText("","Use this option to have the script delete completed items. Consider filtering the`nview to only include a status of NOTIFIED or IN PROCESS.")
eRxGUI.AddText("","Notifications for RxChange requests of at least a day old will be resent as well.")
eRxGUI.AddText("y+8","NOTE: Submission can take a very long time to complete and WQM will appear`nunresponsive. Consider doing other work while this completes.")
eRxGUI.AddLink("y+10", DocLinkText)
eRxGUI.AddButton("ys+25","Surescripts Cancel").OnEvent("Click", SSCancel)
eRxGUI.AddButton("y+50","Delete Completed").OnEvent("Click", eRxCleanup)

; Function call will reload script, hiding GUI.
eRxGUI.OnEvent("Close", (*) => eRxGUI.Destroy())
eRxGUI.OnEvent("Escape", (*) => eRxGUI.Destroy())
eRxGUI.Show "w700"

; Wait for GUI window to close before returning call.
WinWaitClose(eRxGUI.Hwnd)
; Return

; Update variables and do next thing based on GUI input.
SSCancel(*){
eRxGUI.Destroy()
eRxPath := "CancelRx"
Return
}

; Update variables and do next thing based on GUI input.
eRxCleanup(*){
eRxGUI.Destroy()
eRxPath := "DeleteRow"
Return
}
}
}
#HotIf

/************************************************************ Alt+Right Mouse Button  ***************************************************************
; This hotkey will be used for RxQuery.
****************************************************************************************************************************************************/

!RButton::{
KeyWait "Alt", "T2"

; Save clipboard to restore
PriorClip := ClipboardAll()

; If we are in MMR, we'll click right here; if in WQM/CM/DM, we'll highlight the row prior to copy
MouseGetPos &xPosMouse, &yPosMouse, &Win
If WinActive("Monitor.*Remote")
    MouseClick "left", 20, yPosMouse

Else if WinActive("PharmNet: Retail")
    MouseClick "left"

Else if !WinActive("ahk_group MSFT")
    Exit

; Clear clipboard and copy highlighted text
A_Clipboard := ""

; If we're on a Power BI report (SSRS = Report Server), navigate context menu
If WinActive("Power BI",,"Report Server")
    {
    MouseClick "Right"
    Send "{Right}"
    Send "{Enter}"
    }
Else Send "^c"

If !ClipWait(2) OR A_Clipboard = ""
    {
    Tooltip "Copy failed, please try again."
    SetTimer () => ToolTip(), -3000
    A_Clipboard := PriorClip
    Exit
    }

; If we've found an Rx Number, we'll do this branch
If A_Clipboard ~= "\d\d\d\d-\d\d\d\d\d\d\d\d" > 0
    {
    If WinExist("Rx Fill History:")
        WinClose("Rx Fill History")

; Restore = 0 since clipboard is used later in script
    RxNumber := SubStr(A_Clipboard,RegExMatch(A_Clipboard,"\d{4}-\d{8}"),13)

; IF we're in a monitor application, move mouse back where it was, then run query based on context + clipboard
    If WinActive("Monitor.*Remote$")
        {
        MouseMove(xPosMouse,yPosMouse)
        If !WinActive("Work Queue Monitor")
            RxQuery(RxNumber,"RxFillHistory",,xPosMouse,yPosMouse,Win)

; 'Duplicate' kickbacks from CMOP have a special query
        Else if RegExMatch(A_Clipboard,"i)\t(024:DUP|030:DUP|009:DUPLICATE RX FROM ANOTHER VAMC:\t)")
            RxQuery(RxNumber,"CMOPDupes")

        Else
            {
            Switch
                {
                Case RegExMatch(A_Clipboard,"\tAcknowledged\t") : Stage := "Acknowledged"
                Case RegExMatch(A_Clipboard,"\tFulfilled\t") : Stage := "Fulfilled"
                Case RegExMatch(A_Clipboard,"\tPackaged\t") : Stage := "Packaged"
                Case RegExMatch(A_Clipboard,"\tPending Transmission\t") : Stage := "Pending Transmission"
                Default: Stage := "Other"
                }
            Status := RegExMatch(A_Clipboard,"\tSuccess\t") ? "Success" : "Fail"
            RxQuery(RxNumber,"RxFillHistory",,xPosMouse,yPosMouse,Win,Stage,Status)
            }
        }
; If we're in MMR, we run the fill history query
    Else if WinActive("PharmNet: Retail")
        RxQuery(RxNumber,"RxFillHistory",,xPosMouse,yPosMouse,Win)

    Else if WinActive("ahk_group MSFT")
        RxQuery(RxNumber,"RxFillHistory")
    }

; If copy did not include an Rx Number value, and we're in MMR
; Check to see if we have copied a prescriber name
; Use '/'' to exclude date columns, # to exclude Sig, prescriber will have alpha and comma
Else if NOT A_Clipboard ~= "[#/]" AND A_Clipboard ~= "i)^[a-z]+,.[a-z]+" AND WinActive("PharmNet: Retail")
    {
; Clean clip for query
    Prescriber := RegExReplace(A_Clipboard,"[\r\n\d]")

    Query := Format("
                (
                SELECT NameFullFormatted
                    , CASE WHEN PositionTask LIKE 'VA%' AND EMAIL LIKE '%@VA.GOV%' AND ActiveIndicator = 1
                        THEN PositionTask
                        WHEN ExternalIndicator = 1 OR PositionTask IS NULL OR PositionTask LIKE '%VIEW ONLY%'
                        THEN 'External Provider'
                        ELSE PositionTask END AS 'UserType'
                    , CONCAT(IIF(NPI IS NULL,'No NPI found','NPI: ' + NPI), IIF(NPI IS NOT NULL AND (NPIStatus != 'Active' OR NPIInactiveDate < GETDATE()),' (Inactive)','')) AS 'NPI'
                    , CONCAT('Physician: ',IIF(PhysicianIndicator = 1,'Y','N')) AS 'Physician'
                    , CONCAT('Active: ',IIF(ActiveIndicator = 1,'Y','N' + CAST(CAST(ActiveStatusDateTime AS Date) AS VARCHAR))) AS 'ActiveProfile'
                FROM LSV.D05_VISN20INFORM.StaffMill SP
                WHERE SP.PersonStaffSID IN (SELECT Distinct OrderProviderPersonStaffSID
                            FROM LSV.D05_VISN20INFORM.RxMill_Prescription RX
                            WHERE Rx.OrderStatus = 'ORDERED'
                                AND Rx.RxSta3n IS NOT NULL) -- Limit to active VA Rxs to limit duplicate results
                    AND NameFullFormatted = '{1}'
                )", Prescriber)

    Results := LSVQuery(Query,1)
    ; Check results and prep message for user
    If !Results
        Results := "No results found for " Prescriber

    ; If carriage return exists prior to last character, we have multiple results
    Else if InStr(Results,"`r",,-2) > 0
        Results := "Search produced multiple results!`n`n" StrReplace(StrReplace(RegExReplace(Results,"\r$"),"`r","`n`n"),"`t","`n")

    Else Results := StrReplace(StrReplace(RegExReplace(Results,"\r$"),"`r","`n`n"),"`t","`n")

    MsgBox(Results,"Prescriber Overview","4096")
    }

Else if A_Clipboard ~= ", \d\d\d\d\d\d\d\d\d\d\d," AND WinActive("PharmNet: Retail")
    {
    Loop Parse A_Clipboard, ","
        {
        If A_LoopField ~= "\d\d\d\d\d\d\d\d\d\d\d"
            {
            NDC := RegExReplace(A_LoopField,"\D")
            Break
            }
        }
    If IsSet(NDC)
        RxItem(NDC, "NDC")
    }

/*
; Pending C-Rx matching is close but we would need to search for Order Catalog instead of Dispense Drug
; Also would need to handle variations between order catalog (e.g. oral tablet) and dispense drug (e.g. tablet)
Else if A_Clipboard ~= "i)^[a-z]"
    AND WinActive("PharmNet: Retail")
    {
    If A_Clipboard ~= "i)See Instructions"
        {
    ; If ordered sig is 'See Instructions' we can parse from there
        Item := RegExReplace(A_Clipboard,"i) See Instructions.*\R?")
        }
    Else
        {
    ; Check for duration and remove if exists ( [0-9] [duration unit]...)
        Item := RegExReplace(A_Clipboard,"\s\d+\s(day|dose|hr|minute|times|week|month).*")
    ; Then strip the dose/route/frequency
    ; I couldn't get the RegEx to work using a starting position of 0 or -1 (end of string), so used https://regex101.com/
    ; Key is the negative lookahead (?![pattern]) to ensure the pattern doesn't exist in the rest of the string
        Item := RegExReplace(Item,"\s\d*\D(?!.*\s\d*[,.-]?\d+).*")
        }
    MsgBox Item
    }
*/
A_Clipboard := PriorClip
}

/************************************************************ Alt+Left Mouse Button  ***************************************************************
; This hotkey primarily performs lookup actions by capturing a search term from an application and using that term in another application or website.
; Hotkey
****************************************************************************************************************************************************/

; PowerChart with Patient open.
#HotIf WinActive("Opened by")
!LButton::
{
; Wait for alt to release and block keyboard temporarily while executing and mouse movement while executing.
BlockControl(1000)
KeyWait "Alt", "T2"
WinGetPos , , &PCw, , "Opened by"
MouseGetPos &xPosMouse, &yPosMouse

Patient := PCPatient()

; If mouse is above name, first try to extract MRN from WinTitle.
If xPosMouse <= 300
    AND yPosMouse >= 100
    AND yPosMouse <= 240
    AND Patient["MRN"] != ""
    {
    MouseOver := "Name"
    Term := "MRN"
    ID := Patient["MRN"]
    }
; If mouse is over encounter, note that and move on.
Else if xPosMouse >= (0.3 * PCw)
    AND xPosMouse <= (0.53 * PCw)
    AND yPosMouse >= 100
    AND yPosMouse <= 240
    {
    MouseOver := "Encounter"
    }

Else MouseOver := ""

; ID is only expected to be set if we clicked above name field and MRN was captured.
; If not, then click to trigger a Custom Information window (encounter list or patient demographics).
If !IsSet(ID) OR NOT ID
    {
    MouseClick "left"
    WinWaitNotActive("Opened by",,2)
    Try WinGetPos , , &W, &H, "Custom Information:"

    If !IsSet(W) OR !IsSet(H)
        Exit

; If 'narrow' Custom Information: window loads, look for MRN field.
; Double click to highlight MRN.
    If ColorFilter = 2
        {
        Send "#^c" ;Disable color filter.
        Sleep 500
        }
    If W / H < 1.6
        AND ImageSearch(&X, &Y, 15, 220, 60, 265, "*75 HBITMAP:*" ImagePCCustomInfoMRN)
        {
        MouseClick "left", X+100, Y+8, 2
        A_Clipboard := ""
        Send "^c"
        Clipwait(1)
        ID := StrReplace(A_Clipboard,"`r`n")
        Term := "MRN"
        }
    If ColorFilter = 2
        Send "#^c" ;Enable color filter.
; If Custom Information loaded, and we didn't find MRN, and we think mouse was over encounter list.
; Double click to highlight FIN.
    If MouseOver = "Encounter" AND (W / H >= 1.6 OR !X OR !Y)
        {
        Term := "FIN"
        MouseClick "left", 470, 397, 2
        A_Clipboard := ""
        Send "^c"
        Clipwait(1)
        ID := StrReplace(A_Clipboard,"`r`n")
        }
; If we can't find the MRN field, see if we can capture DoD ID via query.
    If MouseOver != "Encounter" AND (W / H >= 1.6 OR !X OR !Y)
        {
        PCPatientDODID(Patient)
        Term := "DODID"
        ID := Patient["DODID"]
        }
    }

If !IsSet(ID) OR NOT ID
    {
    MsgBox("✔ Phone`n✔ Wallet`n❌ Keys`n`nUnable to capture ID for search. Please try again.","PharmNet","T5 4096")
    Exit
    }

; Close custom information window.
Try
    {
    WinClose("Custom Information:")
    }


If Term = "FIN"
    {
    Try MedManagerPatientLookup(ID,Term)
    Catch error as err
        {
        LogError(err)
        PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\sad trombone.mp3")
        MsgBox Format("Hotkey appears to have failed, but FIN: {1} is copied. Try to paste into Med Manager search field or running hotkey again. Sorry about that.", ID),"Hotkey Failed","4096 T5"
        WinWaitClose("Hotkey Failed")
        PlaySound()
        }
    }

Else
    {
    Try MMRPatientLookup(ID,Term)
    Catch Error as err
        {
        LogError(err)
        PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\sad trombone.mp3")
        MsgBox(Format("Hotkey appears to have failed, but {2}: {1} is copied. Try to paste into MMR search field or running hotkey again. Sorry about that.", ID, Term),"Confirmation Failed","4096 T5")
        WinWaitClose("Confirmation Failed")
        PlaySound()
        }
    Else if Term = "RxNumber" AND RxSearchStop = 4
        {
        If MMRFindRx(ID)
            Send "+{Space}"
        }
    Else if Term = "RxNumber" AND RxSearchStop = 3
        MMRAmbProfileSort("RxNumber")
    Else if Term = "RxNumber" AND RxSearchStop = 2
        MMRAmbProfileSort("OrderSentence")
    }
}

; Patient List or Message Center
#HotIf WinActive("PowerChart Organizer for")
!LButton::
{
; Wait for alt to release and block keyboard and mouse movement temporarily
BlockControl(1000)
KeyWait "Alt", "T2"

; Click where mouse is and attempt copy
; If we are on patient list, we should be able to extract FIN
; If not, custom information window should load and we can go that route
MouseClick "left"
A_Clipboard := ""
Send "^c"
Clipwait(1)

; FIN values won't always be 8 or 9 characters long but this works for now.
; Truncating at 10 characters will strip trailing tab if 9 digit FIN.
; Users will have best results if MRN column is not displayed.
ID := StrReplace(SubStr(A_Clipboard,RegExMatch(A_Clipboard,"\t\d\d\d\d\d\d\d\d\d?\t"),10),"`t")

; If we have a FIN, update Term variable.
If ID
    Term := "FIN"
; If not and active window is not custom information, exit
Else if !WinActive("Custom Information:")
    Exit
If ColorFilter = 2
        {
        Send "#^c" ;Disable color filter.
        Sleep 500
        }
If !ID AND ImageSearch(&X, &Y, 15, 220, 60, 265, "*75 HBITMAP:*" ImagePCCustomInfoMRN)
    ; Double click to highlight MRN.
    {
    MouseClick "left", X+100, Y+8, 2
    A_Clipboard := ""
    Send "^c"
    Clipwait(2)
    ID := StrReplace(A_Clipboard,"`r`n")
    Term := "MRN"
    }
If ColorFilter = 2
    Send "#^c" ;Enable color filter.

; Close custom information window.
Try
    {
    WinClose("Custom Information:")
    }

If !IsSet(ID) OR NOT ID
    {
    MsgBox("✔ Phone`n✔ Wallet`n❌ Keys`n`nUnable to capture ID for search. Please try again.","PharmNet","T5 4096")
    Exit
    }

If Term = "FIN"
    {
    Try MedManagerPatientLookup(ID,Term)
    Catch error as err
        {
        LogError(err)
        PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\sad trombone.mp3")
        MsgBox Format("Hotkey appears to have failed, but FIN: {1} is copied. Try to paste into Med Manager search field or running hotkey again. Sorry about that.", ID),"Hotkey Failed","4096 T5"
        WinWaitClose("Hotkey Failed")
        PlaySound()
        }
    }
Else
    {
    Try MMRPatientLookup(ID,Term)
    Catch Error as err
        {
        LogError(err)
        PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\sad trombone.mp3")
        MsgBox(Format("Hotkey appears to have failed, but {2}: {1} is copied. Try to paste into MMR search field or running hotkey again. Sorry about that.", ID, Term),"Confirmation Failed","4096 T5")
        WinWaitClose("Confirmation Failed")
        PlaySound()
        }
    }
}

; Begins with PharmNet: ... Monitor - \\Remote (Claims, Dispense, Work Queue)
#HotIf WinActive("i)^PharmNet: .* Monitor - \Q\\Remote\E")
!LButton::
{
; Wait for alt to release and block keyboard temporarily while executing and mouse movement while executing
; Record current mouse position
Term := "Rx"
BlockControl(1000)
KeyWait "Alt", "T2"
; Script is more consistent if we capture mouse coordinates and pass explicit to click
MouseGetPos &xPosMouse, &yPosMouse

If WinActive("Dispense")
    {
    WinActivate("Dispense Monitor")
; If user on fill/check/pick up, we can extract rx number from row
    MouseClick "left", xPosMouse, yPosMouse
    ID := StringExtract(Term,0.5,0)
; If that fails, double click on highlighted row to open details
; Wait up to 1 second for window to load
; If clicking in empty space, window never launches so exit
    If !IsSet(ID) OR NOT ID
        {
        MouseClick "left", xPosMouse, yPosMouse,2
        If WinWaitActive("View Details", ,1)
            {
; Cursor is Rx Number field when window opens, so we can shift+end to highlight/copy
            Loop 3
                {
                Send "+{End}"
                A_Clipboard := ""
                Send "^c"
                If Clipwait(0.4)
                    Break
                Send "{Home}"
                }
            ID := StrReplace(A_Clipboard,"`r`n")
            WinClose("View Details")
            }
        }
    }

; Called from Claims Monitor
Else if WinActive("Claims")
    {
    WinActivate("Claims Monitor")
; Clicking to highlight row
;    MouseClick "left", 15, yPosMouse, 1
    ID := CMRow(yPosMouse)["RxNumber"]
    }

; Called from Work Queue Monitor
Else if WinActive("Work Queue")
    {
    WinActivate("Work Queue Monitor")
; If user clicked in Rx Number cell, we can use the cell value
    MouseClick "left", xPosMouse, yPosMouse
    A_Clipboard := ""
    Send "^c"
    Clipwait(1)
    ID := StrReplace(A_Clipboard,"`r`n")
; If not, highlight row to try and extract Rx Number from there
    If !IsInteger(StrReplace(ID,"-"))
        ID := WQMRow(yPosMouse)["RxNumber"]
    }

If !IsSet(ID) OR NOT ID
    {
    MsgBox("✔ Phone`n✔ Wallet`n❌ Keys`n`nUnable to capture ID for search. Please try again.","PharmNet","T5 4096")
    Return
    }

Try
    {
    MMRPatientLookup(ID,Term)
    If RxSearchStop = 4
        If MMRFindRx(ID)
            Send "+{Space}"
    Else if RxSearchStop = 3
        MMRAmbProfileSort("RxNumber")
    Else if RxSearchStop = 2
        MMRAmbProfileSort("OrderSentence")
    }
Catch Error as err
    {
    LogError(err)
    PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\sad trombone.mp3")
    MsgBox(Format("Hotkey appears to have failed, but {2}: {1} is copied. Try to paste into MMR search field or running hotkey again. Sorry about that.", ID, Term),"Confirmation Failed","4096 T5")
    WinWaitClose("Confirmation Failed")
    PlaySound()
    }
}

; Called in MMR with patient loaded
; If below banner, add inactive orders
; If in banner, we will either run PharmacyPatient or launch PowerChart (button disabled for students)
#HotIf WinActive("- PharmNet: Ret.*- \Q\\Remote\E")
!LButton::
{
MouseGetPos , &y
If y > 125
    {
    BlockInput "On"
    BlockInput "MouseMove"
    KeyWait "Alt", "T2"

    If MMRAddInactive(0.5)
        MMRAmbProfileSort("OrderSentence")
    BlockInput "Off"
    BlockInput "MouseMove" "Off"

    Exit
    }

; If called in banner, we can try to extract an identifier (MRN, FIN, DoD ID)
; Capturing DoD ID directly should be slightly quicker for Pharmacy Patient lookup
; MRN and FIN are PC search terms and could possibly be used for other use cases
; Since we'll be using the clipboard, we'll try to restore the clipboard when we're done
Else
    {
    PriorClip := ClipboardAll()
    A_Clipboard := ""
; Double clicking ID field in banner enacts copy
    MouseClick "left", , , 2
    ClipWait(0.25)

; Check if we have a value in clipboard and parse if we do
    If RegExMatch(A_Clipboard,"MRN: |DOD ID:")
        {
; MRN and DoD ID should be strictly numeric
; FIN may have "LPE"
; Term should have only alpha, no spaces, etc.
        Term := SubStr(A_Clipboard, 1, InStr(A_Clipboard,":") - 1)
        ID := RegExReplace(A_Clipboard, "MRN: |DOD ID:", "")
        }
    }

; If ID is FIN or MRN, we'll open PowerChart
If IsSet(Term)
    AND Term = "MRN"
    {
    ; Confirm PowerChart is open
    If !WinExist("PowerChart Organizer for")
        {
        ToolTip Term " (" ID ") was copied, but PowerChart may not be open. Please open PowerChart and try again."
        SetTimer () => ToolTip(), -3000
        A_Clipboard := PriorClip
        Exit
        }
    ; If yes, tooltip to ACK user
    ToolTip Term " (" ID ") copied, initiating PowerChart search."
    SetTimer () => ToolTip(), -3000
    ; Open PowerChart and initiate search
    WinActivate("PowerChart Organizer for")
    WinGetPos(,,&W,,"PowerChart Organizer for")
    ; ImageSearch will set X and Y = "" if not found
    If ColorFilter = 2
        {
        Send "#^c" ;Disable color filter.
        Sleep 500
        }
    Loop 20
        {
        Sleep 200
        ImageSearch(&X, &Y, W-60, 110, W, 250, "*75 HBITMAP:*" ImagePCQuickSearch)
        } until X != "" AND Y != ""
    If ColorFilter = 2
        Send "#^c" ;Enable color filter.
    If X = ""
        OR Y = ""
        {
        A_Clipboard := ID
        MsgBox("Unable to locate Quick Search button in PowerChart. MRN: " ID " is copied, please proceed with search manually.","What Once was the P Button","4096 T5")
        Exit
        }
    If Term = "MRN"
        AND ID != ""
        {
    ; Trigger dropdown and short wait for dropdown to load
        MouseClick "left", X+20, Y+12
        Sleep 300
    ; MRN is first on list (keyboard navigation not allowed from PowerChart Organizer window)
        MouseClick "left", X, Y+35, 1
        Sleep 300
    ; Click into search window, send text, wait for text to send
        MouseClick "left", X-20, Y+12
        Send ID
        Sleep 500
        Send "{Enter}"
        }
    Else if Term = "FIN"
        AND ID != ""
        {
        MouseClick "left", X+20, Y+12
        Sleep 300
        MouseClick "left", X, Y+80
        Sleep 300
        MouseClick "left", X-20, Y+12
        Send ID
        Sleep 250
        Send "{Enter}"
        }
    A_Clipboard := PriorClip
    Exit
    }

; If DoD ID, set default ID value to skip LSV query
Else if IsSet(Term)
    AND IsSet(ID)
    {
    DefID := ID
    }

; User clicked in banner but we did not extract the necessary ID value
; So we will instead query to try and capture DoD ID using patient name
Else
    {
    MMRPatient(&PatientName, &DOB, &Query)
    PatientSearch := LSVQuery(Query,0,5,3)

    DefID := ""

    If PatientSearch
        {
        PatientArray := StrSplit(RegExReplace(PatientSearch,"\R"),"|")
; String result often results in additional blank row; check if exists and delete if yes.
; Preference would be to figure out why blank row is created and prevent but this works for now.
        If PatientArray.Get(-1) = ""
            {
            PatientArray.RemoveAt(-1)
            }

        If PatientArray.Length = 1
            {
            PatientRow := StrSplit(PatientArray[1],"`t")
            DefID := PatientRow.Has(2) ? PatientRow[2] : ""
            }
        }
    }
Eligibility(,,DefID)
A_Clipboard := PriorClip
}

; Prescription window [Inquire, New Order, Modify, Refill, Bill Only] for Patient
; CheckHP if called over OK
; CMOP product search if called over Product button
#HotIf WinActive("for Patient:.*- \Q\\Remote\E")
!LButton::
{
BlockControl(1000)
KeyWait "Alt", "T2"
; Capture mouse position relative to window so we know what to do
MouseGetPos &xPosMouse, &yPosMouse, &MouseWin
WinGetPos , , &wWin, &hWin, MouseWin

; If we're hovering over OK, run CheckHP
If !WinActive("i)^Inquire")
    AND hWin - yPosMouse > 9
    AND hWin - yPosMouse < 35
    AND wWin - xPosMouse > 107
    AND wWin - xPosMouse < 199
    {
    If CheckHPOn = 1
        RxCheckHP()
    MouseClick "left", xPosMouse, yPosMouse
    }

; If we're hovering over the title, check eligibility
Else if yPosMouse < 25
    {
    MMRPatient(&PatientName, &DOB, &Query)
    PatientSearch := LSVQuery(Query,0,5,3)

    DefID := ""

    If PatientSearch
        {
        PatientArray := StrSplit(RegExReplace(PatientSearch,"\R"),"|")
; String result often results in additional blank row; check if exists and delete if yes.
; Preference would be to figure out why blank row is created and prevent but this works for now.
        If PatientArray.Get(-1) = ""
            {
            PatientArray.RemoveAt(-1)
            }

        If PatientArray.Length = 1
            {
            PatientRow := StrSplit(PatientArray[1],"`t")
            DefID := PatientRow.Has(2) ? PatientRow[2] : ""
            }
        }
    Eligibility(,,DefID)
    }

; If neither of the two above, else condition is to click and check if product window appears
; If window does appear, search CMOP product line
; If not, exit
Else
    {
    Try
        {
        MouseClick
        If WinWaitActive("Product", ,1)
            {
            NDC := StringExtract("NDC",1,1)
            WinClose("^Product - \Q\\Remote\E")
            }
        If !IsSet(NDC) OR NOT NDC
            Return
        If RxItemPosition ~= "[1-3]"
            RxItem(NDC,"NDC")
        Else
            Run Format('msedge.exe "https://vaww.pbi.cdw.va.gov/PBIRS/Pages/ReportViewer.aspx?/GPE/CMOP_DART/DART/ProductLine&Product={1}"', NDC)
        }
    Catch error as err
        {
        LogError(err)
        }
    }
}

; Open CMOP product line for highlighted row
#HotIf WinActive("Product Selection")
!LButton::
{
BlockControl(1000)
KeyWait "Alt", "T2"

If RxItemPosition ~= "[1-3]"
    {
    NDC := StringExtract("NDC",1,1)
    If IsSet(NDC) AND NDC
        RxItem(NDC,"NDC")
    }
Else
    {
    CMOPID := StringExtract("CMOPID",1,1)
    If IsSet(CMOPID) AND CMOPID
        Run Format('msedge.exe "https://vaww.pbi.cdw.va.gov/PBIRS/Pages/ReportViewer.aspx?/GPE/CMOP_DART/DART/ProductLine&Product={1}"', CMOPID)
    }
}

; PowerBI reports. Must be placed above MSFT group to avoid qualifying on browser
#HotIf WinActive("Power BI",,"Report Server")
!LButton::
{
BlockControl(1000)
KeyWait "Alt", "T2"
A_Clipboard := ""
MouseClick "Right"
Sleep 100
Send "{Right}"
Send "{Enter}"
ClipWait(1)
; Will equal 1 if rx number with hyphen is beginning of clipboard value
; If so, extract and remove any trailing hidden characters
If RegExMatch(A_Clipboard,"\d\d\d\d-\d\d\d\d\d\d\d\d") = 1
    {
    ID := SubStr(A_Clipboard,RegExMatch(A_Clipboard,"\d\d\d\d-\d\d\d\d\d\d\d\d"),13)
    Term := "Rx"
    }
; If copied value is strictly numeric, treat it like MRN
; MRN length can vary so strip trailing characters if exist
Else if RegExMatch(A_Clipboard,"\D") = 0
    {
    ID := RegExReplace(A_Clipboard,"\D")
    Term := "MRN"
    }

If !IsSet(ID) OR NOT ID
    {
    MsgBox("✔ Phone`n✔ Wallet`n❌ Keys`n`nUnable to capture ID for search. Please try again.","PharmNet","T5 4096")
    Exit
    }

Try
    {
    MMRPatientLookup(ID,Term)
    If RxSearchStop = 4
        If MMRFindRx(ID)
            Send "+{Space}"
    Else if RxSearchStop = 3
        MMRAmbProfileSort("RxNumber")
    Else if RxSearchStop = 2
        MMRAmbProfileSort("OrderSentence")
    }
Catch Error as err
    {
    LogError(err)
    PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\sad trombone.mp3")
    MsgBox(Format("Hotkey appears to have failed, but {2}: {1} is copied. Try to paste into MMR search field or running hotkey again. Sorry about that.", ID, Term),"Confirmation Failed","4096 T5")
    WinWaitClose("Confirmation Failed")
    PlaySound()
    }
}

; Called from Teams, Outlook, Excel, or browser
#HotIf WinActive("ahk_group MSFT")
!LButton::
{
BlockControl(1000)
KeyWait "Alt", "T2"
; If Excel, click wherever cursor is
; for Teams or Outlook, user must have already highlighted the value
If WinActive("- Excel")
    {
    MouseClick "left"
    }
A_Clipboard := ""
Send "^c"
Clipwait(1)
; Will equal 1 if rx number with hyphen is beginning of clipboard value
; If matches, extract and clean
If RegExMatch(A_Clipboard,"\d\d\d\d-\d\d\d\d\d\d\d\d") = 1
    {
    ID := SubStr(A_Clipboard,RegExMatch(A_Clipboard,"\d\d\d\d-\d\d\d\d\d\d\d\d"),13)
    Term := "Rx"
    }
; Rx currently (some day this will change) will only have value 2-6 in 5th position
; Rx prefix should only begin with 0-5 for foreseeable future
; If so, extract and clean
; Guessing that this is rx number but could be MRN
Else if RegExMatch(A_Clipboard,"[0-5]\d\d\d[2-6]\d\d\d\d\d\d\d") = 1 AND StrLen(RegExReplace(A_Clipboard,"\D","")) = 12
    {
    ID := SubStr(A_Clipboard,RegExMatch(A_Clipboard,"\d\d\d\d\d\d\d\d\d\d\d\d"),12)
    Term := "Rx"
    }
; Min MRN length is 10 characters
; MRN length can vary so remove non-digit characters if exist
Else if RegExMatch(A_Clipboard,"\d\d\d\d\d\d\d\d\d\d") = 1
    {
    ID := RegExReplace(A_Clipboard,"\D")
    Term := "MRN"
    }
Else
    Exit

If !IsSet(ID) OR NOT ID
    {
    MsgBox("✔ Phone`n✔ Wallet`n❌ Keys`n`nUnable to capture ID for search. Please try again.","PharmNet","T5 4096")
    Exit
    }

Try
    {
    MMRPatientLookup(ID,Term)
    If RxSearchStop = 4
        {
        If MMRFindRx(ID)
            Send "+{Space}"
        }
    Else if RxSearchStop = 3
        MMRAmbProfileSort("RxNumber")
    Else if RxSearchStop = 2
        MMRAmbProfileSort("OrderSentence")
    }
Catch Error as err
    {
    LogError(err)
    PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\sad trombone.mp3")
    MsgBox(Format("Hotkey appears to have failed, but {2}: {1} is copied. Try to paste into MMR search field or running hotkey again. Sorry about that.", ID, Term),"Confirmation Failed","4096 T5")
    WinWaitClose("Confirmation Failed")
    PlaySound()
    }
}

#HotIf WinActive("Suspended Orders Activity Log|Dispense Details")
!LButton::
{
BlockControl(1000)
KeyWait "Alt", "T2"

; Copy row and parse tracking number
A_Clipboard := ""
Send "^c"
Clipwait(0.5)
Tracking := StrReplace(SubStr(A_Clipboard,InStr(A_Clipboard,"`t",,,-1)),"`t")

; "Tracking info unavailable" should not trigger search, so instead make it big!
If InStr(StrUpper(Tracking),"TRACKING")
    {
    MaxWin("A")
    Exit
    }
; If called on Ack row, load CMOP Cerner Rx Lookup report
Else if InStr(A_Clipboard,"Acknowledged")
    {
    ALWin := WinGetTitle("Suspended Orders Activity Log - <")
    RxNumber := SubStr(ALWin,InStr(ALWin,"<")+1,13)
    Run Format('msedge.exe "https://vaww.pbi.cdw.va.gov/PBI_RS/report/GPE/CMOP_Analytics/DART/Cerner%20Rx%20Lookup?SearchTerm={1}"', RxNumber)
    }
; USPS starts with 9
Else if InStr(A_Clipboard,"USPS") OR SubStr(Tracking,1,1) = "9"
    Run Format('msedge.exe "https://tools.usps.com/go/TrackConfirmAction?qtc_tLabels1={1}"', Tracking)
; UPS starts with 1Z
Else if InStr(A_Clipboard,"UPS") OR SubStr(Tracking,1,2) = "1Z"
    Run Format('msedge.exe "https://www.ups.com/track?track=yes&trackNums={1}&loc=en_US&requester=ST/trackdetails"', Tracking)
; Fedex is only other carrier so far, seems to typically start with 7
Else if InStr(A_Clipboard,"FEDEX")
    Run Format('msedge.exe "https://www.fedex.com/fedextrack/?cntry_code=us&trknbr={1}"', Tracking)
Else
    {
; Make it big!
    MaxWin("A")
    Exit
    }
A_Clipboard := Tracking
}

#HotIf WinActive("E-Prescription")
;NPI/diagnosis code lookup.
!RButton::
{
KeyWait "Alt", "T2"
MouseClick('L',,,3) ;Triple-click to select entire row, which may have multiple codes.
Dx := Trim(CaptureRow())
If !Dx
    Exit
Else if RegExMatch(Dx,"^1\d\d\d\d\d\d\d\d\d$") ;NPI?
    {
    Run "https://npiregistry.cms.hhs.gov/provider-view/" Dx
    Exit
    }
codes := StrSplit(Dx, " ") ;May contain a mix of ICD-10 and SNOMED-CT.
matched := 0
For code IN codes
    {
    If !code OR RegExMatch(code,"[^A-Z0-9.]")
        Continue
    Else if RegExMatch(code,"^[A-Z]\d+\.?\w*$")
        {
        Run "https://icdlist.com/icd-10/" code
        matched += 1
        }
    Else if RegExMatch(code,"^\d+$")
        If ValidateSNOMED(code)
            {
            ;Run "https://lookup.snomedtools.org/" code
            Run "https://bioportal.bioontology.org/ontologies/SNOMEDCT?p=classes&conceptid=" code
            matched += 1
            }
    }
If !matched
    {
    ToolTip "No valid diagnosis codes found.",,,3
    SetTimer () => ToolTip(,,,3), -3000
    }
}

;Add provider to X-Rx.
!LButton::
{
If !WinExist("New Order")
    Exit

BlockControl(1000)
KeyWait "Alt", "T2"

If RegExReplace(WinGetTitle("A")," - \Q\\Remote\E") = "E-Prescription"
    EPCS := false
Else
    EPCS := true

PriorClip := ClipboardAll()
CS := 0
mismatch := 0

Try
    eRx := ERxDetails()
Catch
    {
    ToolTip "Failed to activate E-Rx Prescription Details window. Please try again."
    SetTimer () => ToolTip(), -3000
    Exit
    }
A_Clipboard := PriorClip
Try
    RegExMatch(eRx["Prescriber NPI"],"(1\d\d\d\d\d\d\d\d\d)",&NPI)
Catch
    {
    ToolTip "Failed to copy NPI."
    SetTimer () => ToolTip(), -3000
    Exit
    }
If !NPI[1]
    {
    ToolTip "Failed to copy NPI."
    SetTimer () => ToolTip(), -3000
    Exit
    }
Try ;Even if a provider has DEA, doesn't appear to transmit for uncontrolled substances.
    {
    RegExMatch(eRx["Prescriber DEA"],"([A-Za-z]{2}\d{7})",&DEA)
    CS := DEA[1]
    }

If Sta3n = 668
    {
    Try
        FirstName := Trim(eRx["Prescriber First Name"])
    Catch
        FirstName := ""
    Try
        LastName := Trim(eRx["Prescriber Last Name"])
    Catch
        LastName := ""
    Try
        ZipCode := Trim(eRx["Prescriber Zip Code"])
    Catch
        ZipCode := ""
    Try
        State := Trim(eRx["Prescriber State"])
    Catch
        State := ""
    Try
        City := Trim(eRx["Prescriber City"])
    Catch
        City := ""
    Try
        Street := Trim(eRx["Prescriber Street"])
    Catch
        Street := ""
    
    If CheckLEIE(NPI[1], LastName, FirstName, ZipCode, State, City, Street, true)
        Exit
    }

WinMatch("New Order for Patient:",3)

; Try to find Prescriber Search icon column using ImageSearch.
If ColorFilter = 2
        {
        Send "#^c" ;Disable color filter.
        Sleep 500
        }
Loop 4
    {
    Sleep 250
    ImageSearch(&X, &Y, 630, 215, 700, 275, "*75 HBITMAP:*" ImageMagnifyingGlass)
    } until X != "" AND Y != ""
If ColorFilter = 2
    Send "#^c" ;Enable color filter.

; If image is found, offset so we can click in box.
; If image was not found, we can try clicking where box is expected.
X := X = "" ? 666 : X + 5
Y := Y = "" ? 250 : Y + 5

; Redundant WinActivate because there are a lot of windows during Rx processing.
WinActivate("New Order for Patient")
MouseClick "left", X, Y
WinMatch("Provider Selection",5)
WinGetPos , , , &h, "Provider Selection"

Send "!c" ;Clear
Send "!a" ;Alias
Send NPI[1]
Send "{Enter 2}"

; Provider exist?
If ColorFilter = 2
        {
        Send "#^c" ;Disable color filter.
        Sleep 500
        }
Loop 4
    {
    Sleep 750
    ImageSearch(&X, &Y, 0, h - 73, 700, h, "*75 HBITMAP:*" ImageNoMatch)
    } until (X != "" AND Y != "") OR !WinExist("Provider Selection")
If ColorFilter = 2
    Send "#^c" ;Enable color filter.

If X != "" AND Y != "" AND WinExist("Provider Selection") ;If not, add 'em.
    {
    ; Redundant WinActivate because there are a lot of windows during Rx processing.
    WinActivate("Provider Selection")
    Send "!p" ;New Provider
    WinMatch("Add New Provider",5)
    ;Send "!n" ;NPI
    ;Send NPI[1] ;Not how it looks, can't actually input NPI here.
    Send "!f" ;First
    Sleep 250
    Try
        Send Trim(eRx["Prescriber First Name"])
    Sleep 250
    Send "!l" ;Last
    Sleep 250
    Try
        Send Trim(eRx["Prescriber Last Name"])
    Sleep 250
    Send "!b" ;Business
    Sleep 250
    Try
        Send Trim(eRx["Prescriber Phone"])
    Sleep 250
    Send "!x" ;Fax Business
    Sleep 250
    Try
        Send Trim(eRx["Prescriber Fax"])
    Sleep 250
    Send "{tab}" ;Direct Email Address
    Sleep 250
    Try
        Send Trim(eRx["Prescriber Email"])
    Sleep 250
    Send "!e" ;Edit Address
    WinMatch("Address Details",5)
    Try
        Send Trim(eRx["Prescriber Street"])
    Sleep 250
    Send "{tab 4}"
    Sleep 250
    Try
        Send Trim(eRx["Prescriber City"])
    Sleep 250
    Send "{tab}"
    Sleep 250
    Try
        Send Trim(eRx["Prescriber State"])
    Sleep 250
    Send "{tab}"
    Sleep 250
    Try
        Send Trim(eRx["Prescriber Zip Code"])
    Sleep 250
    Send "{enter}"
    WinMatch("Add New Provider",5)
    Send "{enter}"
    WinMatch("Provider Selection",5)
    Send "{enter}"
    WinMatch("Search Physician",5)
    Send "!p" ;Update Identifiers
    WinMatch("Define Aliases",5)
    If ColorFilter = 2
        {
        Send "#^c" ;Disable color filter.
        Sleep 500
        }
    Loop 3
        {
        Sleep 250
        } until ImageSearch(&Xa, &Ya, 0, 0, 139, 186, "*75 HBITMAP:*" ImageAddAlias)
    If ColorFilter = 2
        Send "#^c" ;Enable color filter.Else
    If Xa = "" OR Ya = ""
        {
        ToolTip "Cannot find 'new prsnl alias' icon. Please add NPI/DEA manually. Exiting script."
        SetTimer () => ToolTip(), -5000
        Exit
        }
    MouseClick "left", Xa+5, Ya+5 ;Add new prsnl alias.
    Sleep 500
    Send NPI[1]
    Send "{tab}" ;Personnel alias pool (Type)
    Sleep 500
    Send "n" ;NPI
    Sleep 500
    If CS
        {
        MouseClick "left", Xa+5, Ya+25 ;Add new prsnl alias.
        Sleep 500
        Send CS
        Sleep 500
        Send "{tab}" ;Personnel alias pool (Type)
        Sleep 500
        Send "d" ;DEA
        Sleep 500
        }
    Send "{tab}"
    Send "!a" ;Apply
    Send "!o" ;OK
    WinMatch("Search Physician",5)
    Send "!o" ;OK
    WinMatch("New Order for Patient",5)
    }
Else
    WinMatch("New Order for Patient",5)
If ColorFilter = 2
    {
    Send "#^c" ;Disable color filter.
    Sleep 500
    }
Loop 3
    {
    Sleep 250
    } until ImageSearch(&Xn, &Yn, 505, 192, 699, 237, "*75 HBITMAP:*" ImageNoNPI) ;Detect missing alias.
If ColorFilter = 2
    Send "#^c" ;Enable color filter.
If Xn != "" OR Yn != ""
    EnteredNPI := false
Else
    EnteredNPI := true

;DEA on file (if needed)?
If !CS AND EnteredNPI
    Exit
If EPCS
    {
    If ColorFilter = 2
        {
        Send "#^c" ;Disable color filter.
        Sleep 500
        }
    Loop 3
        {
        Sleep 250
        } until ImageSearch(&Xc, &Yc, 7, 84, 722, 455, "*75 HBITMAP:*" ImageDEA) ;Find DEA field.
    If ColorFilter = 2
        Send "#^c" ;Enable color filter.
    If Xc = "" OR Yc = ""
        {
        Tooltip "Cannot determine presence of DEA on file.", 881, 32
        SetTimer () => ToolTip(), -3000
        }
    Else
        {
        MouseClick(,Xc+102,Yc+26)
        ;Copy
        A_Clipboard := ""
        Loop 3
            {
            Send "^c"
            If Clipwait(1)
                Break
            }
        If CS AND (A_Clipboard = "" OR A_Clipboard = "(None)")
            {
            Send "!r" ;Prescriber
            WinMatch("Search Physician",5)
            Send "!p" ;Update Identifiers
            WinMatch("Define Aliases",5)
            If ColorFilter = 2
                {
                Send "#^c" ;Disable color filter.
                Sleep 500
                }
            Loop 3
                {
                Sleep 250
                } until ImageSearch(&Xa, &Ya, 0, 0, 139, 186, "*75 HBITMAP:*" ImageAddAlias)
            If ColorFilter = 2
                Send "#^c" ;Enable color filter.Else
            If Xa = "" OR Ya = ""
                {
                ToolTip "Cannot find 'new prsnl alias' icon. Please add NPI/DEA manually. Exiting script."
                SetTimer () => ToolTip(), -5000
                Exit
                }
            MouseClick "left", Xa+5, Ya+5 ;Add new prsnl alias.
            Sleep 500
            Send CS
            Sleep 500
            Send "{tab}" ;Personnel alias pool (Type)
            Sleep 500
            Send "d" ;DEA
            Sleep 500
            If !EnteredNPI ;If needed as well, add NPI while here.
                {
                MouseClick "left", Xa+5, Ya+25 ;Add new prsnl alias.
                Sleep 500
                Send NPI[1]
                Sleep 500
                Send "{tab}" ;Personnel alias pool (Type)
                Sleep 500
                Send "n" ;NPI
                Sleep 500
                EnteredNPI := true
                }
            Send "{tab}"
            Send "!a" ;Apply
            Send "!o" ;OK
            WinMatch("Search Physician",5)
            Send "!o" ;OK
            WinMatch("New Order for Patient",5)
            }
        Else if CS != RegExReplace(A_Clipboard,"\s+")
            mismatch := RegExReplace(A_Clipboard,"\s+")
        A_Clipboard := PriorClip
        }
    }

;Add NPI if needed.
If !EnteredNPI
    {
    Send "!r" ;Prescriber
    WinMatch("Search Physician",5)
    Send "!p" ;Update Identifiers
    WinMatch("Define Aliases",5)
    If ColorFilter = 2
        {
        Send "#^c" ;Disable color filter.
        Sleep 500
        }
    Loop 3
        {
        Sleep 250
        } until ImageSearch(&Xa, &Ya, 0, 0, 139, 186, "*75 HBITMAP:*" ImageAddAlias)
    If ColorFilter = 2
        Send "#^c" ;Enable color filter.Else
    If Xa = "" OR Ya = ""
        {
        ToolTip "Cannot find 'new prsnl alias' icon. Please add NPI manually. Exiting script."
        SetTimer () => ToolTip(), -5000
        Exit
        }
    MouseClick "left", Xa+5, Ya+5 ;Add new prsnl alias.
    Sleep 500
    Send NPI[1]
    Sleep 500
    Send "{tab}" ;Personnel alias pool (Type)
    Sleep 500
    Send "n" ;NPI
    Sleep 500
    Send "{tab}"
    Send "!a" ;Apply
    Send "!o" ;OK
    WinMatch("Search Physician",5)
    Send "!o" ;OK
    WinMatch("New Order for Patient",5)
    }

;Legit?
If mismatch
    {
    PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\bell.wav")
    MsgBox("DEA on file (" mismatch ") does not match DEA from eRx (" CS "). Please ensure profile DEA matches eRx. You may need to inactivate current profile alias and add DEA from eRx.", "Warning: DEA Mismatch", "4096")
    ToolTip "Checking profile DEA: " mismatch, 881, 32
    CheckDEA(mismatch)
    ToolTip "Checking eRx DEA: " CS, 881, 32
    CheckDEA(CS)
    }
Else if CS AND EPCS
    CheckDEA(CS)
}
#HotIf
/*
;Rx tracking search from DRP: "Rx Dispense Status by MRN" Report
#HotIf WinActive("Reporting Portal - \\\\Remote")
!LButton::{
MouseGetPos(&X,&Y)
KeyWait "Alt", "T2"
MouseClick , X, Y, 3
rowText := Trim(CaptureRow())
If RegExMatch(rowText, "(\d{4}-\d{8})", &Rx)
    {
    MMRPatientLookup(Rx[1],"Rx")
    If RxSearchStop = 4
        {
        If MMRFindRx(Rx[1])
            Send "+{Space}"
        }
    Else if RxSearchStop = 3
        MMRAmbProfileSort("RxNumber")
    Else if RxSearchStop = 2
        MMRAmbProfileSort("OrderSentence")
    }
Else if InStr(rowText, "Tracking Number: ")
    {
    TN := RegExReplace(rowText, "^.*Tracking Number: ")
    If InStr(rowText, " USPS,") OR SubStr(TN,1,1) = "9"
        Run Format('msedge.exe "https://tools.usps.com/go/TrackConfirmAction?qtc_tLabels1={1}"', TN)
    Else if InStr(rowText, " UPS,") OR SubStr(TN,1,2) = "1Z"
        Run Format('msedge.exe "https://www.ups.com/track?track=yes&trackNums={1}&loc=en_US&requester=ST/trackdetails"', TN)
    Else ;if InStr(rowText, " FEDEX,")
        Run Format('msedge.exe "https://www.fedex.com/fedextrack/?cntry_code=us&trknbr={1}"', TN)
    }
Else
    MsgBox("✔ Phone`n✔ Wallet`n❌ Keys`n`nUnable to capture ID for search. Please try again.","ID Not Found","T5 4096")
}
#HotIf
*/
#HotIf WinActive("- \Q\\Remote\E")
!LButton::MaxWin("A")
#HotIf

/************************************************************ Ctrl+F (Find) ************************************************************************
; Hotkey is used to allow users to perform basic searches.
****************************************************************************************************************************************************/

#HotIf WinActive("Retail Med.*- \\\\Remote|for Patient:.*- \\\\Remote")
^f::
{
; Wait for ctrl key to release to avoid odd behavior, then open GUI.
KeyWait "Ctrl", "T2"
GUIFind()

GUIFind(*)
{
    FindGUI := Gui()
    FindGUI.Title := "Lost and Found"
    FindGUI.Opt("+AlwaysOnTop")
    FindGUI.AddText(,"Search Type: ")
    FindGUI.AddDropDownList("vType x+5 w185", ["Pharmacy Item","Order Catalog","Rx Summary","Rx Fill History"]).OnEvent("Change", UpdateDefaultTerm)
    FindGUI.AddText("xm section","Search Term: ")
    FindGUI.AddEdit("vTerm x+5 w185").OnEvent("Focus", SelectAll)
    FindGUI.AddButton("xm section default x100 w80","Go!").OnEvent("click", submit)

    FindGUI.Show("w280")
    FindGUI.OnEvent("Close", (*) => FindGUI.Destroy())
    FindGUI.OnEvent("Escape", (*) => FindGUI.Destroy())

; This function will update the field to display search terms applicable to the selected search Type
; Use position index from drop down list and RxItemOn setting to determine what to display.
UpdateDefaultTerm(GuiCtrlObj, Info){
Switch
{
    case GuiCtrlObj.Value = 1 : Term := "CMOP ID, NDC, or Name"
    case GuiCtrlObj.Value = 2 : Term := "CMOP ID or NDC"
    case GuiCtrlObj.Value = 3 : Term := "Rx Number or Order ID"
    case GuiCtrlObj.Value = 4 : Term := "Rx Number"
    default: Term := ""
}
FindGUI['Term'].Text := Term
}

; This function simply highlights the pre-populated text in the search term field for easy replacement.
SelectAll(*){
Send "^a"
}

; Function to save GUI input and initiate search.
Submit(*)
{
Saved := FindGUI.Submit(1)
Type := Saved.Type
Term := Saved.Term

; First define category, so we can define query
; CMOP ID is alpha-alphanumeric-numeric-numeric-numeric (e.g. I0004, XC036)
; NDC can be 9 or 11 digits, with or without hyphens
; NOTE: Pseudo-NDCs can result in 'item mixing' (e.g. 10006-0701 includes psyllium and cholecalciferol products)
; For description, search term must start with 3 consecutive alpha characters, and must have at least 5 non-white space characters total
; NDCs less than 11 digits will have a '%' appended
; Description will always have a '%' appended
; Descriptions will have spaces replaced with '%', and a '%' will also be inserted when alpha switches to numeric (e.g. hctz12.5 -> hctz%12.5)
If Type = "Pharmacy Item"
    {
    Switch
        {
        case RegExMatch(Term,"i)^[A-Z]\w\d\d\d$") > 0 : Cat := "CMOPID"
        case RegExMatch(Term,"\d\d\d\d\d-?\d\d\d\d-?\d?\d?") > 0 : Cat := "NDC"
        case RegExMatch(Term,"^[A-Za-z][A-Za-z][A-Za-z]\w\w|^[A-Za-z][A-Za-z][A-Za-z].\w|^[A-Za-z][A-Za-z][A-Za-z]\w.\w") > 0 : Cat := "Description"
        }
    If IsSet(Cat)
        {
        If Cat = "NDC"
            {
            Term := StrLen(Term) < 11 ? Term . '%' : Term
            }
        Else if Cat = "Description"
            {
; Would need a loop if we want to allow multi-drug+strengh searches (e.g. hctz + lisinopril)
            Term := RegExMatch(Term,"i)[a-z][0-9]",&Match) > 0 ? SubStr(Term,1,Match.Pos) . ' ' . SubStr(Term,Match.Pos+1,StrLen(Term)) : Term
            Term := '%' . StrReplace(Term," ","%") . '%'
            }
        RxItem(Term,Cat)
        }
    Else
        {
        If StrLen(Term) < 5
            {
            MsgBox("Search term must be at least 5 characters long, not including spaces. The first 3 characters must be a letter.`n`nPlease try again with an updated search term.","Short Term: " Term,"T10 4096")
            }
        Else
            {
            MsgBox("Unable to perform search for " Term ". Please try again.","Uncertain Terms","T5 4096")
            }
        }
    }
; Order Catalog allows search by CMOP ID or NDC, simliar to Pharmacy Item above
Else if Type = "Order Catalog"
    {
    Searched := Term
    Switch
        {
        case RegExMatch(Term,"i)^[A-Z]\w\d\d\d$") > 0 : Cat := "CMOPID"
        case RegExMatch(Term,"\d\d\d\d\d-?\d\d\d\d-?\d?\d?") > 0 : Cat := "NDC"
        }
    If IsSet(Cat)
        {
        If Cat = "NDC"
            {
            Term := StrLen(Term) < 11 ? Term . '%' : Term
            }
        RxItem(Term,Cat,"OrderCatalog")
        }
    Else
        {
        MsgBox("Unable to perform search for " Searched ". Please try again.","Uncertain Terms","T5 4096")
        }
    }
; Rx Summary allows search by Rx Number (w/w/o hyphen) or person order id (currently 10 or 11 digits)
Else if Type = "Rx Summary"
    {
    If RegExMatch(Term,"^\d\d\d\d-?\d\d\d\d\d\d\d\d$|^\d\d\d\d\d\d\d\d\d\d\d?$")
        {
        Term := RegExMatch(Term,"^\d\d\d\d\d\d\d\d\d\d\d\d") = 1 ? SubStr(Term,1,4) . "-" . SubStr(Term,5,8) : Term
        RxQuery(Term,"RxSummary")
        }
    Else
        {
        MsgBox("Unable to perform search for " Term ". Please try again.","Uncertain Terms","T5 4096")
        }
    }
; Rx Fill History allows search by Rx Number (w/w/o hyphen)
Else if Type = "Rx Fill History"
    {
    If RegExMatch(Term,"^\d\d\d\d-?\d\d\d\d\d\d\d\d$")
        {
        Term := RegExMatch(Term,"^\d\d\d\d\d\d\d\d\d\d\d\d") = 1 ? SubStr(Term,1,4) . "-" . SubStr(Term,5,8) : Term
        RxQuery(Term,"RxFillHistory")
        }
    Else
        {
        MsgBox("Unable to perform search for " Term ". Please try again.","Uncertain Terms","T5 4096")
        }
    }
}
}
}
#HotIf

/************************************************************ Ctrl+Left-Click ****************************************************************************
; Enter CITC PA Consult
*********************************************************************************************************************************************************/

; PowerChart with Patient chart open.
#HotIf WinActive("Opened by") AND MMRCtrlLButton = 1
^LButton::
{
KeyWait "Control", "T2"
VetProC := RegExReplace(WinGetTitle("A")," Opened by.*\Q\\Remote\E") ;Name/MRN from active win, to prevent another open profile from being referenced.
Tooltip "Adding BVE", 881, 32
PCAddBVE(VetProC)
WinWait("i)Assign a Relationship",,2)
If WinExist("i)Assign a Relationship") AND !WinActive("i)Assign a Relationship")
    WinActivate("i)Assign a Relationship")
If WinActive("i)Assign a Relationship")
    WinWaitActive(VetProC)
Tooltip "Ordering Consult", 881, 32
PCNewOrder("Outside Provider",, VetProC, RegExReplace(VetProC, " - \d+"))
Tooltip "Preparing Consult", 881, 32
InitiateNFPA(,,VetProC)
Tooltip
PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\CLICK_Nice.mp3")
}

/************************************************************ Alt+Ctrl+Left-Click ****************************************************************************
; run the Drug Info Search from the New, Modify, and Verify MM screens
*********************************************************************************************************************************************************/

;; Run Infosearch from Verify Order Screen
;;selects the "VA" agency and runs the report
;#HotIf WinActive("^(Verify|Modify|New).*Order - \\\\Remote") RegExMatch
;+!LButton::
;{
;KeyWait "LButton", "T0.5"
;
;A_Clipboard := ""
;Send "^c"
;Clipwait(3)
;
;Infosearch()
;}
;#HotIf

; Extract details from order details to facilitate fax communication.
#HotIf WinActive("E-Prescription Detail")
^LButton::
{
Try
    eRx := ERxDetails()

Catch
    {
    ToolTip "Failed to activate E-Rx Prescription Details window. Please try again."
    SetTimer () => ToolTip(), -3000
    }

global LastXRxCopy := ""
Try
    LastXRxCopy .= "Patient Name: " eRx["Patient First Name"] " " eRx["Patient Last Name"] "`n"
Try
    LastXRxCopy .= "Patient DOB: " eRx["Patient DOB"] "`n"
Try
    LastXRxCopy .= "Item: " eRx["Drug Name"] "`n"
Try
    LastXRxCopy .= "Prescriber Name: " eRx["Prescriber First Name"] " " eRx["Prescriber Last Name"] "`n"
Try
    LastXRxCopy .= "NPI: " eRx["Prescriber NPI"] "`n"
Try
    LastXRxCopy .= "Fax: " eRx["Prescriber Fax"]
A_Clipboard := LastXRxCopy
ToolTip("Copied to clipboard:`n`n" LastXRxCopy)
SetTimer () => ToolTip(), -4000
}
#HotIf

/************************************************************ SPO Toolkit GUI ****************************************************************
; Suite of things that are in development or niche projects for Spokane.
; Will expand to other sites if/when able.
**********************************************************************************************************************************************/

; Named function hotkey for GUIMain button to trigger Toolkit menu.
#HotIf Sta3n = "668"
!+s::
SPOTools(ThisHotkey)
{
GUISPOToolkit()
}
#HotIf

/************************************************************ Fax Mover *********************************************************************************
; Rename and move electronic fax files in network folder
; Initial use case is CITC PADR related
; Additional use cases may include various CFU documents, faxed rxs, faxes other to allow triage of all faxed documents
*********************************************************************************************************************************************************/
#HotIf WinActive("SPO Toolkit")
FaxMover(){
If Sta3n != "668"
    Exit

SourceFolder := "\\va.gov\pa\Spokane-Walla Walla\VHASPO\Fax\Pharmacy"
DestinationFolder := "\\va.gov\pa\Spokane-Walla Walla\VHASPO\Services\Pharmacy"

If !IsSet(FaxFile)
    {
    FaxFile := "", FaxFriendly := ""
    xWin := "x" MonitorMiddle, yWin := "y30"
    }

If WinExist("CITC Fax")
    WinActivate("CITC Fax")

Else GUIFaxMove(xWin " " yWin)

GUIFaxMove(*)
{
FaxMoveGUI := Gui()
FaxMoveGUI.Opt("+AlwaysOnTop")
FaxMoveGUI.Title := "CITC Fax Mover"
FaxMoveGUI.AddButton("w75 Section","&Select File").OnEvent("Click",call_file_select)
FaxMoveGUI.AddButton("x+10 w75", "&Move File").OnEvent("Click", Submit)
FaxMoveGUI.AddButton("x+10 w75", "&Reset").OnEvent("Click", Reset)
FaxMoveGUI.AddText("xm","Patient Last Name: ")
FaxMoveGUI.AddEdit("r0.8 vLastName Section w150 x+4 yp-2")
FaxMoveGUI.AddText("xm","Patient First Name: ")
FaxMoveGUI.AddEdit("r0.8 vFirstName Section w150 x+5 yp-2")
FaxMoveGUI.AddText("xm","File Selected: ")
FaxMoveGUI.AddText("x+2 r1 vFF w220")
FaxMoveGUI.AddText("xm", "Destination Folder: ")
FaxMoveGUI.AddText("x+2 r1 vDD w220")
FaxMoveGUI.OnEvent("Close", (*) => FaxMoveGUI.Destroy())
FaxMoveGUI.OnEvent("Escape", (*) => FaxMoveGUI.Destroy())
FaxMoveGUI.Show(xWin yWin)

; Function to save GUI input
Submit(*)
{
    FaxMoveGUI.GetClientPos(&x, &y)
    xWin := x, yWin := y
    Saved := FaxMoveGUI.Submit(0)
    If FaxFile = ""
        {
        MsgBox("No file selected. Please try again.","Please Select Fax","4096 T3")
        Return
        }
    If Saved.LastName = "" OR Saved.FirstName = ""
        {
        MsgBox("Please populate patient name fields.","Patient Name Missing","4096 T3")
        Return
        }

    Loop 5
        {
        Try
            {
            WinActivate(FaxFriendly)
            Send "^w"
            }
        } until !WinExist(FaxFriendly)

    NewFile := DestinationFolder "\" Saved.LastName "_" Saved.FirstName ".pdf"

    Loop
        {
        Try
            {
            FileMove FaxFile, NewFile
            ToolTip "File renamed and moved to: " NewFile
            SetTimer () => ToolTip(), -3000
            Break
            }
        Catch error as err
            {
            If A_Index < 4
                {
                Sleep 1000
                NewFile := StrReplace(NewFile,"_",A_Index "_")
                If A_Index = 1
                    MsgBox "Selected folder may already contain PDF named " Saved.LastName "_" Saved.FirstName ". Will attempt to save as " NewFile " instead.", "Duplicate File Name", "4096 T4"
                Else
                    {
                    LastAnnex := A_Index - 1
                    MsgBox "Selected folder may already contain PDF named " Saved.LastName LastAnnex "_" Saved.FirstName ". Will attempt to save as " NewFile " instead.", "Duplicate File Name", "4096 T4"
                    }
                Continue
                }
            Else
                {
                PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\sad trombone.mp3")
                MsgBox(FaxFriendly " could not be moved. Ensure you have access to the destination folder. If selected directory already contains similarly named files, manual rename/move may be necessary.","Failed to Move File","4096 T5")
                WinWaitClose("Failed to Move File")
                PlaySound()
                }
            LogError(err)
            Return
            }
        }
    FaxFile := "", FaxFriendly := "", NewFile := ""
    FaxMoveGUI['FF'].Value := FaxFriendly
    FaxMoveGUI['LastName'].Value := ""
    FaxMoveGUI['FirstName'].Value := ""
    Return
}

Reset(*)
{
    Loop 5
        {
        Try
            {
            WinActivate(FaxFriendly)
            Send "^w"
            }
        } until !WinExist(FaxFriendly)

    FaxFile := "", FaxFriendly := ""
    FaxMoveGUI['FF'].Value := FaxFriendly
    FaxMoveGUI['LastName'].Value := ""
    FaxMoveGUI['FirstName'].Value := ""
	Return
}

call_file_select(*)
{
    FaxMoveGUI.Opt("OwnDialogs")
    Try
        {
        FaxFile := FileSelect("3", SourceFolder,, "*.pdf")
        If !FaxFile
            Return
        }
    Catch
        {
        ToolTip("Folder not found.")
        SetTimer () => ToolTip(), -2500
        Return
        }
    Run FaxFile
    FaxFriendly := SubStr(FaxFile,InStr(FaxFile,"\",,-1)+1)
    FaxMoveGUI['FF'].Value := FaxFriendly
    SourceFolder := RegExReplace(FaxFile, "\\.*\.pdf$")
    DestinationFolder := RegExReplace(DirSelect(DestinationFolder,3,"Select Destination Folder for Fax Files"), "\\$")
    If !DestinationFolder
        DestinationFolder := "\\va.gov\pa\Spokane-Walla Walla\VHASPO\Services\Pharmacy"
    FaxMoveGUI['DD'].Value := RegExReplace(DestinationFolder,"^\\\\.*va\.gov.*\\Services","S:")
    Return
}
}
}

/************************************************************ CFU Templates *****************************************************************************
; Rename and move electronic fax files in network folder
; Initial use case is CITC PADR related
; Additional use cases may include various CFU documents, faxed rxs, faxes other to allow triage of all faxed documents
*********************************************************************************************************************************************************/

PADRTemplates(){
If !Sta3n = 668
    Exit

Pharmacy := "MANN-GRANDSTAFF VA MEDICAL CENTER`n4815 NORTH ASSEMBLY STREET`nPHARMACY SERVICE, 119`nSPOKANE, WASHINGTON  99205`nFAX: (509)434-7111`nPHONE: (509)434-7022"
TemplateDirectory := String(StrReplace(A_ScriptFullPath, "AHK\Oracle Cerner Repository for Automating Pharmacy-VA.ahk") "PADR\Non-Formulary Templates")
If !IsSet(LastXRxCopy)
    global LastXRxCopy := ""

GUIFaxTemplates()

GUIFaxTemplates(*)
{
FaxTemplateGUI := Gui()
FaxTemplateGUI.Title := "CCN Fax Machine - " Sta3n
FaxTemplateGUI.AddText("", "Pharmacy: " SubStr(Pharmacy,1,InStr(Pharmacy,"`n")-1)).OnEvent("Click",call_pharm_detail)
FaxTemplateGUI.AddText("y+5", "Template Directory: ")
FaxTemplateGUI.AddText("x+2 r1 vTD w400").OnEvent("DoubleClick",call_file_select)
FaxTemplateGUI['TD'].Value := StrReplace(TemplateDirectory, "\\R01spohsm01.r01.med.va.gov\Services","S:")
FaxTemplateGUI.AddButton("w150 h25 xm+5 Section","Change Template Directory").OnEvent("Click",call_file_select)
FaxTemplateGUI.AddButton("w150 h25 x+20","Copy Rx Details from X-Rx").OnEvent("Click",call_detail_copy)
FaxTemplateGUI.AddButton("w150 h25 x+20","Copy Last Rx to Clipboard").OnEvent("Click",call_last_rx)
FaxTemplateGUI.OnEvent("Close", (*) => FaxTemplateGUI.Destroy())
FaxTemplateGUI.OnEvent("Escape", (*) => FaxTemplateGUI.Destroy())
If A_UserName ~= "i)(YURCIE|YEAMAC|DICKHE)"
    {
    ContextMenu := Menu()
    ContextMenu.Add "&Delete File", ContextSearch
    }

; Create the ListView with two columns, Name and Size:
LV := FaxTemplateGUI.Add("ListView", "xm r10 w500 Grid -multi", ["Name"])

; Notify the script whenever the user double clicks a row:
LV.OnEvent("DoubleClick", LV_DoubleClick)

/*
ShowContextMenu enables the right-click file delete option.
This is useful for 'lock files' but may not be a desirable option
Re-enable as needed for file clean up or if requested
*/

LV.OnEvent("ContextMenu", ShowContextMenu)

; Gather a list of file names from a folder and put them into the ListView:
Loop Files, TemplateDirectory "\*.*"
    LV.Add("Vis", A_LoopFileName)

FaxTemplateGUI.Show

LV_DoubleClick(LV, RowNumber){
; Get the text from the row's first field.
TemplateFile := TemplateDirectory "\" LV.GetText(RowNumber)
Run TemplateFile
FaxTemplateGUI.Minimize()
Return
}

; In response to right-click or Apps key.
ShowContextMenu(LV, Item, IsRightClick, X, Y){
CoordMode "Mouse", "Client"
; Show the menu at the provided coordinates, X and Y.  These should be used
; because they provide correct coordinates even if the user pressed the Apps key:
MouseGetPos &xCM, &yCM
ContextMenu.Show(xCM, yCM)
}

; Controlled by ShowContextMenu above
ContextSearch(ItemName, *){
FocusedRowNumber := LV.GetNext(0, "F")  ; Find the focused row.
if not FocusedRowNumber  ; No row is focused.
    return
Try
    {
    FileRecycle TemplateDirectory "\" LV.GetText(FocusedRowNumber)
    LV.Delete(FocusedRowNumber)
    }
Catch error as err
    {
    PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\sad trombone.mp3")
    MsgBox("Unable to delete file. Please try again later.","Delete Failed","4096 T5")
    LogError(err)
    WinWaitClose("Delete Failed")
    PlaySound()
    }
Return
}

call_file_select(*){
FaxTemplateGUI.Opt("OwnDialogs")
NewDir := FileSelect("D",TemplateDirectory,"Select Folder - CCN Fax Machine Template Directory")
If NewDir != ""
    TemplateDirectory := NewDir
FaxTemplateGUI['TD'].Value := StrReplace(TemplateDirectory, "\\R01spohsm01.r01.med.va.gov\Services","S:")
Return
}

call_detail_copy(*){
FaxTemplateGUI.Minimize()
Loop 2
    {
    If !WinWaitActive("E-Prescription",,15)
        {
        OKGo := MsgBox("When an E-Prescription window becomes active, click 'OK' to proceed.","Waiting for Active E-Prescription Window","4096 OC")
        If OKGo != "OK"
            Return
        }
    } until WinActive("E-Prescription")
global LastXRxCopy := ERxDetails()
ToolTip("Copied to clipboard:`n`n" LastXRxCopy)
SetTimer () => ToolTip(), -2500
}

call_last_rx(*){
If LastXRxCopy != ""
    {
    A_Clipboard := LastXRxCopy
    ToolTip("Copied to clipboard:`n`n" LastXRxCopy)
    }
Else
    ToolTip("Last Rx not found")
SetTimer () => ToolTip(), -2500
}

call_pharm_detail(*){
ToolTip(Pharmacy)
SetTimer () => ToolTip(), -4000
}
}
}

/************************************************************ Local CS Batch Screener *******************************************************************
; Scrolls WQM batch (intended local CS items) and notifies user if suspend until date may be earlier than desired.
; Default anchor date/suspend until date calculation has a 10 day buffer so 'external' refills often suspend early
; Script will copy rx, query LSV, result if 'early'
; User can update rx via resuspend in WQM or refill in MMR
*********************************************************************************************************************************************************/

LocalBatch(){
; Check for open WQM and if exists, activate and maximize.
If !WinExist("Work Queue Monitor")
    {
    MsgBox("Please open Work Queue Monitor then try again.","Unable to Find WQM","4096 T5")
    Exit
    }

If !WinActive("Work Queue Monitor")
    WinActivate("Work Queue Monitor")
If !WinGetMinMax("Work Queue Monitor")
    MaxWin("Work Queue Monitor")
WinGetPos , , , &hWQM, "Work Queue Monitor"

; Initialize variables.
LastRx := " "
ForReview := ""
PullCount := 0
LastRow := hWQM - 118
BSMode := 0
BSDueDate := A_Now

; Use GUI (WinTitle = Local Batch Due Check) to capture date
; Closing GUI returns without setting BSMode
GUIBatchScreen(&BSMode)

If !IsSet(BSMode)
    Exit

Tooltip "Please wait, script is executing", 881, 32, 1

Loop
    {
    WinActivate("PharmNet: Work Queue Monitor")
; If first pass, yPosMouse = first row (window is maximized)
    If A_Index = 1
        yPosMouse := 153
; Else if yPosMouse can be increased within bounds, do so
    Else if yPosMouse + 18 <= LastRow
        yPosMouse += 18
; Else leave yPosMouse unchanged but change row using down arrow
; Sleep is to avoid sending double click and opening detail window accidentally
    Else
        {
        Send "{Down}"
        Sleep 400
        }

    Rx := WQMRow(yPosMouse)

;    MouseClick "left", 20, yPosMouse, 1
;    StringExtract(&RxNumber,"Rx",2,0)

;    ReqMethod := RegExMatch(A_Clipboard,"\tInternal\t") > 0 ? "Internal" : "External"
;    SuccessFail := RegExMatch(A_Clipboard,"\tSuccess\t") > 0 ? 1 : 0

; If end of list or copy failed, break loop
    If !IsObject(Rx)
        OR LastRx = Rx["RxNumber"]
        OR A_Clipboard = "Pull Local"
        {
        Break
        }
; Don't query if in failed status or internal (assume set accurately when internal)
    Else if Rx["Status"] = "Fail"
        OR Rx["RequestMethod"] = "Internal"
        {
        LastRx := Rx["RxNumber"]
        Sleep 250
        Continue
        }
; Query if external + success
    Result := RxQuery(Rx["RxNumber"],"LocalMailScreen",BSDueDate)

; Function returns 1 if fill is confirmed due
; Update action to pull local if user has option turned on
    If Result = 1 AND BSMode = 1
        {
        Send "{Home}"
        Send "p"
        Send "{Enter}"
        LastRx := Rx["RxNumber"]
        PullCount += 1
        }
; Function returns concated string (rx number \tab due date or 'not found' \new line)
; Concatenate results for summary message
    Else if Result != 1
        ForReview .=  Result

    LastRx := Rx["RxNumber"]
    }

; Complete any populated actions now
; We will only submit if a pull local action exists, hence if followed by loop (instead of allowing Loop 0 to handle itself)

If PullCount > 0
    {
    WinActivate("Work Queue Monitor")

    Send "!s"

    Loop PullCount
        {
        WinMatch("for Patient:",15)
        If CheckHPOn = 1
            RxCheckHP()
        Send "!o"
        If !WinWaitClose("for Patient",,15)
            MsgBox("Unable to confirm pull action completed. Exiting script.","This is Weird","4096")
        }

    }

; If early fills were found, provide results for review
If ForReview != ""
    {
    A_Clipboard := ForReview
    MsgBox("These Rxs may be early:`n`n" ForReview "`nResults have also been loaded to clipboard if you would like to paste somewhere.","Batch Screen Results","4096")
    }
Else
    {
    PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\CLICK_Nice.mp3")
    MsgBox("Script is complete","Batch Screen Complete","4096 T5")
    }

Tooltip

GUIBatchScreen(*)
{
BatchScreenGUI := Gui()
BatchScreenGUI.Title := "Local Batch Due Check"
BatchScreenGUI.Opt("+AlwaysOnTop")
BatchScreenGUI.AddText("","Fill due date (on or before):")
BatchScreenGUI.AddDateTime("x+8 w120 vDate")
BatchScreenGUI.AddText("xm","Perform Pull Local When Due: ")
BatchScreenGUI.AddCheckBox("x+8 vBSMode Checked" BSMode)
BatchScreenGUI.AddText("xm y+10","The script will begin at row 1 after you click start.")
BatchScreenGUI.AddButton("x130 w40 Default","&Start").OnEvent("Click",call_start)
BatchScreenGUI.OnEvent("Close", (*) => BatchScreenGUI.Destroy)
BatchScreenGUI.OnEvent("Escape", (*) => BatchScreenGUI.Destroy)
BatchScreenGUI.Show("w290")
; Wait for GUI window to close before returning call
WinWaitClose("Local Batch Due Check")

call_start(*){
Saved := BatchScreenGUI.Submit(1)
BSDueDate := Saved.Date, BSMode := Saved.BSMode
Return BSMode
}
}
}
#HotIf

/************************************************************ Ctrl + N (Bedrock Nom.) ************************************************************
; Nominate providers for access to prescribe controlled substances.
; Starts on "Welcome" win.
; Loop until all nominated.
; This shouldn't be on pharmacy... for URAC to review.
******************************************************************************************************************************************************/

#HotIf WinActive("ePrescribing Controlled Substance Logical Access Control Setup")
^N::
{
Welcome := WinGetID("A") ;These windows all have the same title, but different ahk_id values.
Tooltip "Nominating providers for " Sta3n ". Please wait.", 881, 32
If ColorFilter = 2
    {
    Send "#^c" ;Disable color filter.
    Sleep 500
    }
Loop 3
    {
    ImageSearch(&xN, &yN, 6, 158, 189, 248, "*75 HBITMAP:*" ImageNom)
    Sleep 250
    } until xN != "" AND yN != ""
If xN != "" AND yN != ""
    MouseClick "left", xN, yN
Else
    {
    If ColorFilter = 2
        Send "#^c" ;Enable color filter.
    Tooltip "Unable to find Nominate button. Please try again."
    SetTimer () => ToolTip(), -3000
    Exit
    }
Loop 3
    {
    ImageSearch(&xB, &yB, 4, 384, 164, 481, "*75 HBITMAP:*" ImageBegin)
    Sleep 250
    } until xB != "" AND yB != ""
If ColorFilter = 2
    Send "#^c" ;Enable color filter.
If xB != "" AND yB != ""
    MouseClick "left", xB, yB
Else
    {
    Tooltip "Unable to find button to Begin. Please try again."
    SetTimer () => ToolTip(), -3000
    Exit
    }
While Welcome = WinGetID("A")
    Sleep 500
Location := WinGetID("A")
Send "{tab 2}"
Send Sta3n ;Station preference saved from home GUI.
Send "{enter}"
If ColorFilter = 2
    {
    Send "#^c" ;Disable color filter.
    Sleep 500
    }
While Location = WinGetID("A")
    {
    Sleep 500
    If ImageSearch(&xR, &yR, 0, 0, 395, 97, "*75 HBITMAP:*" ImageNoResultsFound)
        {
        PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\CLICK_Nice.mp3")
        If ColorFilter = 2
            Send "#^c" ;Enable color filter.
        Tooltip "No providers to nominate for " Sta3n "."
        SetTimer () => ToolTip(), -3000
        Exit
        }
    }
Loop
    {
    Send "{tab}" ;Highlight location.
    Send "!n" ;Next
    While !ImageSearch(&xP, &yP, 407, 122, 824, 206, "*75 HBITMAP:*" ImageNomProviders)
        Sleep 250 ;Wait for next page, since ahk_id doesn't change... at least not always?
    Send "{tab 8}"
    Send "{space}" ;Move all providers from 'Available' to 'Nominate' column.
    Select := WinGetID("A")
    Send "!f" ;Finish
    While Select = WinGetID("A")
        Sleep 500
    Welcome := WinGetID("A")
    MouseClick "left", xN, yN
    Sleep 250
    MouseClick "left", xB, yB
    While Welcome = WinGetID("A")
        Sleep 500
    Location := WinGetID("A")
    Send "{tab 2}"
    Send Sta3n ;Station preference saved from home GUI.
    Send "{enter}"
    While Location = WinGetID("A")
        {
        Sleep 500
        If ImageSearch(&xR, &yR, 0, 0, 395, 97, "*75 HBITMAP:*" ImageNoResultsFound)
            {
            PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\CLICK_Nice.mp3")
            If ColorFilter = 2
                Send "#^c" ;Enable color filter.
            Tooltip "All providers nominated for " Sta3n "."
            SetTimer () => ToolTip(), -3000
            Exit
            }
        }
    }
}
#HotIf

/************************************************************ Alt + Shift + X (Exit) *********************************************************************
; Hotkey will interrupt executing script, release any input blocks that may exist, and exit the script
; Script must be reloaded before hotkey will run again
; Panic button that will hopefully never be used
; ScriptInfo('KeyHistory') will likely include PHI and should not be included, ListVars and ListHotkeys
; don't offer much additional information
*********************************************************************************************************************************************************/

!+x::
{
KeyWait "Alt", "T2"
; Collect lines recently executed
FileAppend(ScriptInfo("ListLines"), A_Desktop "\ErrorLog.txt")
PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\The World.mp3") ; Stop time.
BlockInput "Off"
BlockInput "MouseMove" "Off"
; Disable Tooltip and blocks if exist
ToolTip

If BetaTester AND Sta3n = 1012
    {
    PanicChoiceGUI()
    return
    }

Report := MsgBox("Exit command called. If you experienced a failure, please click OK to send a report via email. An error log file has been created and will be attached to the email.`n`nIf everything is fine, click Cancel to resume the script. Have a nice day.","Catastrophic Exit","OC Icon!")
If Report = "OK"
    {
    PlaySound("\\va.gov\pa\Spokane-Walla Walla\VHASPO\Media\Pharmacy\AHK-SFX\Darth Vader NO!.mp3")
    ; In memory of OCRAP's glorious founder, Lewis DeJaegher, whose inbox bravely received these until his retirement.
    SendTo := "VHAPharmNOVA@va.gov"
    If !SendEmail("PharmNOVA Error: NoVa", "Here's what I was doing and what went wrong...", SendTo,,, 2,, A_Desktop "\ErrorLog.txt", true)
        {
        Run("mailto:" SendTo "?subject=PharmNOVA Error: NoVa&body=Here's what I was doing and what went wrong.")
        If WinWaitActive("PharmNOVA Error: NoVa",,5)
            {
            Send "{alt}"
            KeyWait "Alt", "T2"
            Send "h"
            Send "af"
            Send "b"
            If WinWaitActive("Insert File",,5)
                {
                Send A_Desktop "\ErrorLog.txt"
                Send "{Enter}"
                }
            Else
                MsgBox("We're having a rough day here. Please attach the ErrorLog.txt file from your desktop to this email prior to sending. Sorry for the inconvenience.","Mo Hotkeys Mo Problems","4096")
            }
        Else
            MsgBox("We're having a rough day here. Can you please attach the ErrorLog.txt file from your desktop to this email prior to sending? Sorry for the inconvenience.","Mo Hotkeys Mo Problems","4096")
        }
    Run "https://forms.office.com/g/c21vpV88Jf"
    FileDelete A_Desktop "\ErrorLog.txt"
    Reload
    }
Else
    {
    PlaySound()
    FileDelete A_Desktop "\ErrorLog.txt"
    LastChance := MsgBox("Previous script will resume after you click OK. If you do not want the script to resume, please exit or reload the script.")
    If LastChance != "OK"
        Exit
    }
}

/************************************************************ Alt + Shift + H (Help/Home) ************************************************************
; Hotkey that opens main GUI.
******************************************************************************************************************************************************/

!+h::
{
If WinExist("Pharmacy National Optimization for Veterans by Autohotkey - User Menu")
    WinClose("Pharmacy National Optimization for Veterans by Autohotkey - User Menu")

GUIMain()
}

/************************************************************ Alt + Shift + T (Hotstrings) ************************************************************
; Hotkey that opens hotstrings.
******************************************************************************************************************************************************/

!+t::
Hotstrings(hk){
If EnableHotstring
    {
    If !HotstringToS
        CheckHotstringToS()
    HotstringGUI()
    }
}

/************************************************************ Alt + Shift + I (InfoSearch) ************************************************************
; Hotkey that opens InfoSearch.
******************************************************************************************************************************************************/

!+i::
Infosearch(*){
InfoSearchGUI()
}

/************************************************************ Alt + Shift + D (SendFax) ************************************************************
; Send fax through Outlook via ECFax.
****************************************************************************************************************************************************/

!+w::
SendFax(*){
ECFax()
}

/************************************************************ Alt + Shift + R (PREVENT) ************************************************************
; PREVENT calc.
******************************************************************************************************************************************************/
!+r::
PREVENT(*){
PREVENTGUI()
}

/************************************************************ Alt + Shift + R (CrCl) ************************************************************
; CrCl calc.
******************************************************************************************************************************************************/
#HotIf BetaTester AND Sta3n = 1011
!+c::
CrCl(*){
CrClGUI()
}
#HotIf

/************************************************************ Alt+Ctrl+Left-Click ****************************************************************************
; run the Drug Info Search from the New, Modify, and Verify MM screens
***********************************************************************************************************************************************************/

;; Run Infosearch from Verify Order Screen
;;selects the "VA" agency and runs the report
;#HotIf WinActive("^(Verify|Modify|New).*Order - \\\\Remote") RegExMatch
;+!LButton::
;{
;KeyWait "LButton", "T0.5"
;
;A_Clipboard := ""
;Send "^c"
;Clipwait(3)
;
;Infosearch()
;}
;#HotIf

; ============================================================
; Local Drug Message Automation – Cerner (Citrix) – AHK v2
; ============================================================
; Purpose:
;   Detect when a user selects a product in Product Selection,
;   clicks OK, capture the selected row, map it to a Local Drug
;   File (LDF) message, and display that message during the
;   New Order workflow.
; ============================================================

#HotIf WinActive("Product Selection - \\\\Remote") AND !WinExist("E-Prescription Change Request") AND !Busy AND !WinExist("New Order.*\\\\Remote") AND IsSet(LDFMessage) AND LDFMessage
LButton::
!o::
Enter::
NumpadEnter::
Tab::
^RButton::
{
global Busy := true
If (WinGetMinMax("A") = 1 AND PixelGetColor(76, 36) = 0xFFFFFF) OR (!WinGetMinMax("A") AND PixelGetColor(78, 44) = 0xFFFFFF) ;Do nothing if on "Settings" tab.
        {
        If A_ThisHotkey = "LButton"
            MouseClick    ; Pass click through - not OK button.
        Else if A_ThisHotkey = "Tab"
            Send "{Tab}"
        Else if InStr(A_ThisHotkey, "Enter")
            Send "{Enter}"
        global Busy := false
        Exit
        }
If A_ThisHotkey = "Tab"
    {
    ToolTip "WARNING: Tab key may cause LDF messaging to fail. Please use alternative clicks/keystrokes to select product or close window.", 0, 33
    SetTimer () => ToolTip(), -5000
    ;Send "{tab}"
    global Busy := false
    Exit
    }

MouseGetPos(&mx, &my)
WinGetPos(, , &winW, &winH, "A")

If A_ThisHotkey = "^RButton"
    {
    If NOT ((WinGetMinMax("A") = 1 AND my < 173 AND my > 151 AND mx < 482 AND mx > 11) OR (!WinGetMinMax("A") AND my < 181 AND my > 159 AND mx < 483 AND mx > 12)) ;Search box.
        {
        global Busy := false
        Exit
        }
    Else if !LSVAccess OR Sta3n != 668 ;Gotta query, local prototype.
        {
        If !LSVAccess AND Sta3n = 668
            {
            ToolTip "This hotkey requires LSV access. Please ask your pharmacy ADPAC for assistance."
            SetTimer () => ToolTip(), -3000
            }
        global Busy := false
        Exit
        }
    ;Highlight text.
    MouseClick("L", mx, my)
    Send "^a" ;all
    search := Trim(CaptureRow())
    If !RegExMatch(search, "\w+") ;Copy
        {
        If search
            {
            ToolTip "Invalid search. Please type generic name and try again."
            SetTimer () => ToolTip(), -3000
            }
        global Busy := false
        Exit
        }
    QueryNDF := Format("SELECT DISTINCT LabelDescription, ItemID, CMOPID, VAFormularyStatus, CopayTier FROM LSV.D05_VISN20INFORM.Product_Lookup_AHK WHERE LabelDescription LIKE '%{1}%'", RegExReplace(search, "\s+", "%"))
    Try
        NDF := LSVQuery(QueryNDF,,5,3)
    Catch
        {
        ToolTip "Query failed. Please try again."
        SetTimer () => ToolTip(), -3000
        global Busy := false
        Exit
        }
    If !RegExMatch(NDF, "\w")
        {
        ToolTip "No matching product found."
        SetTimer () => ToolTip(), -3000
        NewItemRequest()
        global Busy := false
        Exit
        }
    options := StrSplit(NDF, "`r")
    RequestFlexGUI()
    }
    
; For LButton, verify click is within OK button region.
Else if A_ThisHotkey = "LButton" {
    rightOff  := winW - mx
    bottomOff := winH - my

    ; OK button region (offsets from bottom-right).
    OK_RIGHT_MIN  := 110
    OK_RIGHT_MAX  := 201
    OK_BOTTOM_MIN := 39
    OK_BOTTOM_MAX := 65
    OK_FUDGE      := 10
    
    If NOT (rightOff  >= (OK_RIGHT_MIN  - OK_FUDGE)
        AND rightOff  <= (OK_RIGHT_MAX  + OK_FUDGE)
        AND bottomOff >= (OK_BOTTOM_MIN - OK_FUDGE)
        AND bottomOff <= (OK_BOTTOM_MAX + OK_FUDGE))
        {
        If (my >= 0 AND my < 31 AND mx >= 0 AND  mx < (winW - 135)) ;Move win.
            OR (my < 5 AND my >= 0) OR (my > (winH - 5) AND my <= winH) OR (mx < 5 AND mx >= 0) OR (mx > (winW - 5) AND mx <= winW) ;Resize win.
            {
            KeyWait "LButton"
            CoordMode "Mouse", "Client"
            MouseGetPos(&mx2, &my2)
            Try MouseClickDrag("L",mx,my,mx2,my2)
            CoordMode "Mouse", "Window"
            global Busy := false
            Exit
            }
        Else if (WinGetMinMax("A") = 1 AND my > 209 AND my < (winH - 90) AND mx < (winW - 14) AND mx > (winW - 32)) OR (!WinGetMinMax("A") AND my > 217 AND my < (winH - 91) AND mx < (winW - 15) AND mx > (winW - 33)) ;Vertical scroll bar.
            {
            KeyWait "LButton"
            MouseGetPos(&mx2, &my2)
            Try MouseClickDrag("L",mx,my,mx,my2)
            global Busy := false
            Exit
            }
        Else if (WinGetMinMax("A") = 1 AND my > 192 AND my < 219 AND mx > 13 AND mx < (winW - 31)) OR (!WinGetMinMax("A") AND my > 201 AND my < 227 AND mx > 14 AND mx < (winW - 32)) ;Resize col.
            OR (WinGetMinMax("A") = 1 AND my > (winH - 91) AND my < (winH - 73) AND mx > 30 AND mx < (winW - 49)) OR (!WinGetMinMax("A") AND my > (winH - 92) AND my < (winH - 74) AND mx > 31 AND mx < (winW - 50)) ;Horizontal scroll bar.
            OR (WinGetMinMax("A") = 1 AND my < 173 AND my > 151 AND mx < 482 AND mx > 11) OR (!WinGetMinMax("A") AND my < 181 AND my > 159 AND mx < 483 AND mx > 12) ;Search box.
            {
            KeyWait "LButton"
            MouseGetPos(&mx2, &my2)
            Try MouseClickDrag("L",mx,my,mx2,my)
            global Busy := false
            Exit
            }
        Else if (WinGetMinMax("A") = 1 AND my > 218 AND my < (winH - 90) AND mx > 13 AND mx < (winW - 31)) OR (!WinGetMinMax("A") AND my > 226 AND my < (winH - 91) AND mx > 14 AND mx < (winW - 32)) ;Double-click row?
            {
            KeyWait "LButton"
            If !KeyWait("LButton", "D T0.1")
                {
                MouseClick , mx, my    ; Pass click through - not OK button.
                global Busy := false
                Exit
                }
            Try if WinExist("Product Selection - \\\\Remote") AND !WinActive("Product Selection - \\\\Remote")
                WinActivate("Product Selection - \\\\Remote")
            If !WinActive("Product Selection - \\\\Remote") OR WinExist("New Order.*\\\\Remote")
                {
                global Busy := false
                Exit
                }
            MouseGetPos(&mx2, &my2)
            If my2 != my OR mx2 != mx
                {
                ;MouseClick , mx2, my2
                global Busy := false
                Exit
                }
            }
        Else if mx > winW OR mx < 0 OR my < 0 OR my > winH ;Clicking onto a monitor with different resolution can lead mouse to jump.
            {
            CoordMode "Mouse", "Client"
            MouseClick    ; Pass click through - not OK button.
            CoordMode "Mouse", "Window"
            global Busy := false
            Exit
            }
        Else
            {
            MouseClick , mx, my    ; Pass click through - not OK button.
            global Busy := false
            Exit
            }
        }
}

Try if A_ThisHotkey != "^RButton"
    LDF_ProcessFlow()
global Busy := false

;Flexing Request Product Selection
RequestFlexGUI(*){
FlexProductSelection := Gui()
FlexProductSelection.Title := "Flexing Request Product Selection - " search
FlexProductSelection.Opt("+AlwaysOnTop")

If options.Length < 36
    lv := FlexProductSelection.AddListView("xm ym w780 r" options.Length " Grid -multi Sort", ["Label Description","Item ID","CMOP ID","Formulary Status","Copay Tier"])
Else
    lv := FlexProductSelection.AddListView("xm ym w780 r35 Grid -multi Sort", ["Label Description","Item ID","CMOP ID","Formulary Status","Copay Tier"])
fr := FlexProductSelection.AddButton("x+10 w80","&Request Flex")
ni := FlexProductSelection.AddButton("y+10 w80","&New Item")

for option in options
    {
    if !RegExMatch(option, "\w+")
        continue
    columns := StrSplit(option, "`t")
    lv.Add(, columns[1], columns[2], columns[3], columns[4], columns[5])
    }
loop 5
    lv.ModifyCol(A_Index,"AutoHdr")
;lv.Modify(1, "+Select")

;fm := GetFlexMessage(lv)
;gm := FlexProductSelection.AddText("y+10 w500 Left Wrap", fm)

;lv.OnEvent("Change", gm.Text := GetFlexMessage(lv))
lv.OnEvent("DoubleClick", (*) => FlexRequest(lv))
fr.OnEvent("Click", (*) => FlexRequest(lv))
ni.OnEvent("Click", (*) => (FlexProductSelection.Minimize(), NewItemRequest()))

HotIfWinActive("Hotstring Manager")
Hotkey("!r", (*) => FlexRequest(lv), "On")
Hotkey("!n", (*) => NewItemRequest(), "On")
HotIfWinActive() ; Clear context.

FlexProductSelection.OnEvent("Escape", (*) => FlexProductSelection.Destroy())
FlexProductSelection.Show()
}
}
#HotIf

;Ensure eRx change requests are transmitted to the correct practice, with adequate service level.
#HotIf WinActive("E-Prescription Change Request - \d{4}-\d{8} - \\\\Remote")
!a::
{
Send "!v" ;View Rx
WinMatch("E-Prescription - \\\\Remote",3)
eRx := ERxDetails()
WinMatch("E-Prescription - \\\\Remote",3)
Send "!c" ;Close
WinMatch("E-Prescription Change Request - \d{4}-\d{8} - \\\\Remote",3)
Send "!u" ;SureScripts Directory
WinMatch("Surescripts Prescriber Directory - \\\\Remote",6)
provider := SureScriptsAddressMatch(eRx["Prescriber Street"], eRx["Prescriber City"], eRx["Prescriber State"], eRx["Prescriber Zip Code"])
WinMatch("E-Prescription Change Request - \d{4}-\d{8} - \\\\Remote",3)
If InStr(provider["ServiceLevels"], "CHANGE")
    {
    Tooltip "Prescriber's address matched and service level for change exists."
    SetTimer () => ToolTip(), -3000
    }
Else
    MsgBox("OCRAP WARNING: SureScripts transmission expected to fail. Prescriber's address matched, but service level for change does not exist. Consider using ECFax instead, supported by Alt+Shift+W hotkey.","Change Not Configured","4096 T5")
}
#HotIf
