ster Outlook with Growl' for Microsoft Outlook 2011 Mac
By Erik Gustavson (http://eigenspace.org)

Based on the script by Matt Legend Gemmell ( http://mattgemmell.com/ or @mattgemmell on Twitter). Original at http://mattgemmell.com/using-growl-with-microsoft-outlook

Details can be found here: http://eigenspace.org/microsoft-outlook-growl-notifications

License:
Copyright (c) 2011 Erik Gustavson

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

(MIT License, see http://www.opensource.org/licenses/mit-license.php)
*)

-- Register a notification type called "New Mail" with Growl, and enable it.
tell application "GrowlHelperApp"
	set the allNotificationsList to {"New Mail"}
	set the enabledNotificationsList to {"New Mail"}
	register as application ¬
		"Outlook" all notifications allNotificationsList ¬
		default notifications enabledNotificationsList ¬
		icon of application "Microsoft Outlook"
end tell
