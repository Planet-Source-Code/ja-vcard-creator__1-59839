vCard Creator
With this project you can create a vCard file (vcf) and send it to a cellular phone using Bluetooth (Widcom) or save it as a file in ANSI (windows) format or UTF-8
When sending it using BT then the format is UTF-8 so there is no problem with non English characters.If you plan to add this contact in windows address book then you should save it in ANSI
I am using this program for over 5 months with my old Nokia 6600 and my new p910i with no problems
You must have Widcom BT stack, and not Microsoft's
The path of the widcom BT app that sends the vcard file is set form computer to
	C:\Program Files\WIDCOMM\Bluetooth Software\btsendto_explorer.exe
and I beleive in most computers is this. If not change it



TODO
1.	Currently, Send over BT displays a system window to select the target. I 
	am trying to find a way to select it directly from my app
2.	A parser. Open a vcf file
3.	Encode a photo using base64 encoding
4.	

