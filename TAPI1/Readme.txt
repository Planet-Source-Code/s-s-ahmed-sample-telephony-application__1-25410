Author: S.S. Ahmed 
Email: ss_ahmed1@hotmail.com
Date: Jul 21, 2001

This sample application shows you how you can use the TAPI 3.0 functions to create a TAPI application of your own. Microsoft provides more than 100 functions as part of TAPI library. TAPI 2.1 was more difficult to use as it was a set of API functions but TAPI 3.0 which comes as part of Windows 2000 has ActiveX controls that makes the TAPI application creation much easier. The TAPI SDK that can be downloaded from the Microsoft site provide several samples which show the usage of TAPI functions in VB, JAVA and C++. Learning to use different TAPI functions require lot of time investment but once you create some applications using these TAPI functions, it becomes easier for you to use the TAPI functions more efficiently.

Requirements to run this sample code:
======================================

This application uses the Windows Phone Dialer utility that comes with Windows and can be found under the Accessories group. That program will act as the call manager and will handle the dialing of the phone number that is passed to the tapiRequestMakeCall function.

Running the program will invoke the call manager, and the call manager then dials the number entered in the textbox.

Before running the sample code, make sure that the modem is installed in your machine, also, to experiment with other TAPI functions, download the complete TAPI SDK (if its now already available on your computer) from the microsoft site. Several third parties develop tools that makes the creation of Telephony applications much easier. If you are a serious TAPI developer than you must download the third party tools from the internet or dwell into the TAPI SDK yourself to get a thorough understanding of the telephony functions. You can download TAPI software free of cost from the following site:

http://www.microsoft.com/communications/telephony.htm

TAPI2195.exe is the TAPI implementation for Windows 95, and TAPI21NT for Windows NT. Download and install the one that is appropriate for your computer.