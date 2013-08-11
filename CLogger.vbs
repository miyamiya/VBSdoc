' This program is free software; you can redistribute it and/or
' modify it under the terms of the GNU General Public License
' as published by the Free Software Foundation; either version 2
' of the License, or (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

'! Create an error message with hexadecimal error number from the given Err
'! object's properties. Formatted messages will look like "Foo bar (0xDEAD)".
'!
'! Implemented as a global function due to general lack of class methods in
'! VBScript.
'!
'! @param  e   Err object
'! @return Formatted error message consisting of error description and
'!         hexadecimal error number. Empty string if neither error description
'!         nor error number are available.
Public Function FormatErrorMessage(e)
  Dim re : Set re = New RegExp
  re.Global = True
  re.Pattern = "\s+"
  FormatErrorMessage = Trim(Trim(re.Replace(e.Description, " ")) & " (0x" & Hex(e.Number) & ")")
End Function

'! Create an error message with decimal error number from the given Err
'! object's properties. Formatted messages will look like "Foo bar (42)".
'!
'! Implemented as a global function due to general lack of class methods in
'! VBScript.
'!
'! @param  e   Err object
'! @return Formatted error message consisting of error description and
'!         decimal error number. Empty string if neither error description
'!         nor error number are available.
Public Function FormatErrorMessageDec(e)
  Dim re : Set re = New RegExp
  re.Global = True
  re.Pattern = "\s+"
  FormatErrorMessage = Trim(Trim(re.Replace(e.Description, " ")) & " (" & e.Number & ")")
End Function

'! Class for abstract logging to one or more logging facilities. Valid
'! facilities are:
'!
'! - interactive desktop/console
'! - log file
'! - eventlog
'!
'! Note that this class does not do any error handling at all. Taking care of
'! errors is entirely up to the calling script.
'!
'! @author  Ansgar Wiechers <ansgar.wiechers@planetcobalt.net>
'! @date    2011-03-13
'! @version 2.0
Class CLogger
	Private validLogLevels
	Private logToConsoleEnabled
	Private logToFileEnabled
	Private logFileName
	Private logFileHandle
	Private overwriteFile
	Private sep
	Private logToEventlogEnabled
	Private sh
	Private addTimestamp
	Private debugEnabled
	Private vbsDebug

	'! Enable or disable logging to desktop/console. Depending on whether the
	'! script is run via wscript.exe or cscript.exe, the message is either
	'! displayed as a MsgBox() popup or printed to the console. This facility
	'! is enabled by default when the script is run interactively.
	'!
	'! Console output is printed to StdOut for Info and Debug messages, and to
	'! StdErr for Warning and Error messages.
	Public Property Get LogToConsole
		LogToConsole = logToConsoleEnabled
	End Property

	Public Property Let LogToConsole(ByVal enable)
		logToConsoleEnabled = CBool(enable)
	End Property

	'! Indicates whether logging to a file is enabled or disabled. The log file
	'! facility is disabled by default. To enable it, set the LogFile property
	'! to a non-empty string.
	'!
	'! @see #LogFile
	Public Property Get LogToFile
		LogToFile = logToFileEnabled
	End Property

	'! Enable or disable logging to a file by setting or unsetting the log file
	'! name. Logging to a file ie enabled by setting this property to a non-empty
	'! string, and disabled by setting it to an empty string. If the file doesn't
	'! exist, it will be created automatically. By default this facility is
	'! disabled.
	'!
	'! Note that you MUST set the property Overwrite to False BEFORE setting
	'! this property to prevent an existing file from being overwritten!
	'!
	'! @see #Overwrite
	Public Property Get LogFile
		LogFile = logFileName
	End Property

	Public Property Let LogFile(ByVal filename)
		Dim fso, ioMode

		filename = Trim(Replace(filename, vbTab, " "))
		If filename = "" Then
			' Close a previously opened log file.
			If Not logFileHandle Is Nothing Then
				logFileHandle.Close
				Set logFileHandle = Nothing
			End If
			logToFileEnabled = False
		Else
			Set fso = CreateObject("Scripting.FileSystemObject")
			filename = fso.GetAbsolutePathName(filename)
			If logFileName <> filename Then
				' Close a previously opened log file.
				If Not logFileHandle Is Nothing Then logFileHandle.Close

				If overwriteFile Then
					ioMode = 2  ' open for (over)writing
				Else
					ioMode = 8  ' open for appending
				End If

				' Open log file either as ASCII or Unicode, depending on system settings.
				Set logFileHandle = fso.OpenTextFile(filename, ioMode, -2)

				logToFileEnabled = True
			End If
			Set fso = Nothing
		End If

		logFileName = filename
	End Property

	'! Enable or disable overwriting of log files. If disabled, log messages
	'! will be appended to an already existing log file (this is the default).
	'! The property affects only logging to a file and is ignored by all other
	'! facilities.
	'!
	'! Note that changes to this property will not affect already opened log
	'! files until they are re-opened.
	'!
	'! @see #LogFile
	Public Property Get Overwrite
		Overwrite = overwriteFile
	End Property

	Public Property Let Overwrite(ByVal enable)
		overwriteFile = CBool(enable)
	End Property

	'! Separate the fields of log file entries with the given character. The
	'! default is to use tabulators. This property affects only logging to a
	'! file and is ignored by all other facilities.
	'!
	'! @raise  Separator must be a single character (5)
	'! @see http://msdn.microsoft.com/en-us/library/xe43cc8d (VBScript Run-time Errors)
	Public Property Get Separator
		Separator = sep
	End Property

	Public Property Let Separator(ByVal char)
		If Len(char) <> 1 Then
			Err.Raise 5, WScript.ScriptName, "Separator must be a single character."
		Else
			sep = char
		End If
	End Property

	'! Enable or disable logging to the Eventlog. If enabled, messages are
	'! logged to the Application Eventlog. By default this facility is enabled
	'! when the script is run non-interactively, and disabled when the script
	'! is run interactively.
	'!
	'! Logging messages to this facility produces eventlog entries with source
	'! WSH and one of the following IDs:
	'! - Debug:       ID 0 (SUCCESS)
	'! - Error:       ID 1 (ERROR)
	'! - Warning:     ID 2 (WARNING)
	'! - Information: ID 4 (INFORMATION)
	Public Property Get LogToEventlog
		LogToEventlog = logToEventlogEnabled
	End Property

	Public Property Let LogToEventlog(ByVal enable)
		logToEventlogEnabled = CBool(enable)
		If sh Is Nothing And logToEventlogEnabled Then
			Set sh = CreateObject("WScript.Shell")
		ElseIf Not (sh Is Nothing Or logToEventlogEnabled) Then
			Set sh = Nothing
		End If
	End Property

	'! Enable or disable timestamping of log messages. If enabled, the current
	'! date and time is logged with each log message. The default is to not
	'! include timestamps. This property has no effect on Eventlog logging,
	'! because eventlog entries are always timestamped anyway.
	Public Property Get IncludeTimestamp
		IncludeTimestamp = addTimestamp
	End Property

	Public Property Let IncludeTimestamp(ByVal enable)
		addTimestamp = CBool(enable)
	End Property

	'! Enable or disable debug logging. If enabled, debug messages (i.e.
	'! messages passed to the LogDebug() method) are logged to the enabled
	'! facilities. Otherwise debug messages are silently discarded. This
	'! property is disabled by default.
	Public Property Get Debug
		Debug = debugEnabled
	End Property

	Public Property Let Debug(ByVal enable)
		debugEnabled = CBool(enable)
	End Property

	' - Constructor/Destructor ---------------------------------------------------

	'! @brief Constructor.
	'!
	'! Initialize logger objects with default values, i.e. enable console
	'! logging when a script is run interactively or eventlog logging when
	'! it's run non-interactively, etc.
	Private Sub Class_Initialize()
		logToConsoleEnabled = WScript.Interactive

		logToFileEnabled = False
		logFileName = ""
		Set logFileHandle = Nothing
		overwriteFile = False
		sep = vbTab

		logToEventlogEnabled = Not WScript.Interactive

		Set sh = Nothing

		addTimestamp = False
		debugEnabled = False
		vbsDebug = &h0050

		Set validLogLevels = CreateObject("Scripting.Dictionary")
		validLogLevels.Add vbInformation, True
		validLogLevels.Add vbExclamation, True
		validLogLevels.Add vbCritical, True
		validLogLevels.Add vbsDebug, True
	End Sub

	'! @brief Destructor.
	'!
	'! Clean up when a logger object is destroyed, i.e. close file handles, etc.
	Private Sub Class_Terminate()
		If Not logFileHandle Is Nothing Then
			logFileHandle.Close
			Set logFileHandle = Nothing
			logFileName = ""
		End If

		Set sh = Nothing
	End Sub

	' ----------------------------------------------------------------------------

	'! An alias for LogInfo(). This method exists for convenience reasons.
	'!
	'! @param  msg   The message to log.
	'!
	'! @see #LogInfo(msg)
	Public Sub Log(msg)
		LogInfo msg
	End Sub

	'! Log message with log level "Information".
	'!
	'! @param  msg   The message to log.
	Public Sub LogInfo(msg)
		LogMessage msg, vbInformation
	End Sub

	'! Log message with log level "Warning".
	'!
	'! @param  msg   The message to log.
	Public Sub LogWarning(msg)
		LogMessage msg, vbExclamation
	End Sub

	'! Log message with log level "Error".
	'!
	'! @param  msg   The message to log.
	Public Sub LogError(msg)
		LogMessage msg, vbCritical
	End Sub

	'! Log message with log level "Debug". These messages are logged only if
	'! debugging is enabled, otherwise the messages are silently discarded.
	'!
	'! @param  msg   The message to log.
	'!
	'! @see #Debug
	Public Sub LogDebug(msg)
		If debugEnabled Then LogMessage msg, vbsDebug
	End Sub

	'! Log the given message with the given log level to all enabled facilities.
	'!
	'! @param  msg       The message to log.
	'! @param  logLevel  Logging level (Information, Warning, Error, Debug) of the message.
	'!
	'! @raise  Undefined log level (51)
	'! @see http://msdn.microsoft.com/en-us/library/xe43cc8d (VBScript Run-time Errors)
	Private Sub LogMessage(msg, logLevel)
		Dim tstamp, prefix

		If Not validLogLevels.Exists(logLevel) Then Err.Raise 51, _
			WScript.ScriptName, "Undefined log level '" & logLevel & "'."

		tstamp = Now
		prefix = ""

		' Log to facilite "Console". If the script is run with cscript.exe, messages
		' are printed to StdOut or StdErr, depending on log level. If the script is
		' run with wscript.exe, messages are displayed as MsgBox() pop-ups.
		If logToConsoleEnabled Then
			If InStr(LCase(WScript.FullName), "cscript") <> 0 Then
				If addTimestamp Then prefix = tstamp & vbTab
				Select Case logLevel
					Case vbInformation: WScript.StdOut.WriteLine prefix & msg
					Case vbExclamation: WScript.StdErr.WriteLine prefix & "Warning: " & msg
					Case vbCritical:    WScript.StdErr.WriteLine prefix & "Error: " & msg
					Case vbsDebug:      WScript.StdOut.WriteLine prefix & "DEBUG: " & msg
				End Select
			Else
				If addTimestamp Then prefix = tstamp & vbNewLine & vbNewLine
				If logLevel = vbsDebug Then
					MsgBox prefix & msg, vbOKOnly Or vbInformation, WScript.ScriptName & " (Debug)"
				Else
					MsgBox prefix & msg, vbOKOnly Or logLevel, WScript.ScriptName
				End If
			End If
		End If

		' Log to facility "Logfile".
		If logToFileEnabled Then
			If addTimestamp Then prefix = tstamp & sep
			Select Case logLevel
				Case vbInformation: logFileHandle.WriteLine prefix & "INFO" & sep & msg
				Case vbExclamation: logFileHandle.WriteLine prefix & "WARN" & sep & msg
				Case vbCritical:    logFileHandle.WriteLine prefix & "ERROR" & sep & msg
				Case vbsDebug:      logFileHandle.WriteLine prefix & "DEBUG" & sep & msg
			End Select
		End If

		' Log to facility "Eventlog".
		' Timestamps are automatically logged with this facility, so addTimestamp
		' can be ignored.
		If logToEventlogEnabled Then
			Select Case logLevel
				Case vbInformation: sh.LogEvent 4, msg
				Case vbExclamation: sh.LogEvent 2, msg
				Case vbCritical:    sh.LogEvent 1, msg
				Case vbsDebug:      sh.LogEvent 0, "DEBUG: " & msg
			End Select
		End If
	End Sub
End Class
