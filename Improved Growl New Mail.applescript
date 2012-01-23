(*
Improved Growl New Mail, for Microsoft Outlook 2011 Mac
By Erik Gustavson (http://eigensoft.com) 

Based on the script by Matt Legend Gemmell ( http://mattgemmell.com/ or @mattgemmell on Twitter). Original at http://mattgemmell.com/using-growl-with-microsoft-outlook

Details can be found here: http://eigenspace.org/microsoft-outlook-growl-notifications

License:
Copyright (c) 2011 Erik Gustavson

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

(MIT License, see http://www.opensource.org/licenses/mit-license.php)
*)

-- don't try to growl if growl is not running (see http://growl.info/documentation/applescript-support.php)
tell application "System Events"
	set isRunning to (count of (every process whose name is "GrowlHelperApp")) > 0
end tell

if isRunning is true then
	
	-- Get a list of all "current messages" in Outlook.
	tell application "Microsoft Outlook"
		set theMessages to the current messages
		
		-- Loop through the messages.
		repeat with theMsg in theMessages
			
			-- Only Growl about unread messages.
			if is read of theMsg is false then
				
				-- subject
				set mysubject to get the subject of theMsg
				
				-- sender
				try
					set mysender to sender of theMsg
					if name of mysender is "" then
						set mysender to address of mysender
					else
						set mysender to name of mysender
					end if
				on error errmesg number errnumber
					try
						set mysender to address of mysender
					on error errmesg number errnumber
						-- Couldn't get name or email; we'll just say the sender is unknown.
						set mysender to "Unknown sender"
					end try
				end try
				
				-- content, truncated to 30 characters
				try
					set mycontent to plain text content of theMsg
					
					set mycontentlen to (length of mycontent)
					if (mycontentlen > 30) then
						set mycontent to (text 1 thru 29 of mycontent) as string
					end if
					
				on error errmesg number errnumber
					set mycontent to "<No Content>"
				end try
				
				-- growl it!
				tell application "Growl"
					notify with name "New Mail" title mysender description (mysubject & " - " & mycontent) application name "Outlook"
				end tell
			end if
			
		end repeat
	end tell
end if

