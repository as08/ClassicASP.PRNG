<%
	' INSTALLATION:
	'************************************************************************
	
	' Uses the .NET RNGCryptoServiceProvider class 
	' https://docs.microsoft.com/en-us/dotnet/api/system.security.cryptography.rngcryptoserviceprovider?view=netframework-4.7.2
	
	' Make sure you have the lastest .NET Framework installed (tested on .NET Framework 4.7.2)
	
	' Open IIS, go to the applicaton pools and open the application pool being used by your 
	' website. Check the .NET CRL version
	' E.g: v4.0.30319
	
	' Navigate to the CRL folder
	' E.g: C:\Windows\Microsoft.NET\Framework64\v4.0.30319
	
	' Copy over: ClassicASP.PRNG.dll
	
	' Run CMD as administrator
	
	' Change the directory to your CRL folder
	' E.g: cd C:\Windows\Microsoft.NET\Framework64\v4.0.30319
	
	' Run the following command: RegAsm ClassicASP.PRNG.dll /tlb /codebase
	
	
	' PRNG IN VBSCRIPT:
	'************************************************************************
		
	function random_string(ByVal strLen)
		
		Dim PRNG : set PRNG = server.CreateObject("ClassicASP.PRNG")
			
			' Generate a random string from the following characters:
			' abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890
			
			random_string = PRNG.RandomString(strLen)
		
		set PRNG = nothing
		
	end function
	
	function random_sanitized_string(ByVal strLen)
		
		Dim PRNG : set PRNG = server.CreateObject("ClassicASP.PRNG")
			
			' The following characters are ommited from sanitized strings:
			' 0 1 2 5   A E I O U   L N S Z
			' This is to avoid any profanity being randomly generated in
			' public facing random strings
			
			random_sanitized_string = PRNG.RandomString(strLen,"bcdfghjkmpqrtvwxyBCDFGHJKMPQRTVWXY346789")
		
		set PRNG = nothing
		
	end function
	
	function random_password(ByVal strLen)
		
		Dim PRNG : set PRNG = server.CreateObject("ClassicASP.PRNG")
			
			' Include special characters
			
			random_password = PRNG.RandomString(strLen,"abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!#$%&""';=*+-/:<>?@\^_~")
		
		set PRNG = nothing
		
	end function
	
	function random_number(ByVal min, ByVal max)
	
		Dim PRNG : set PRNG = server.CreateObject("ClassicASP.PRNG")
			
			' Generate a random number between a minimum and maximum value
			
			random_number = PRNG.RandomNumber(min,max)
			
		set PRNG = nothing
		
	end function
	
	function random_token(ByVal keyLen, ByVal encoding)
	
		Dim PRNG : set PRNG = server.CreateObject("ClassicASP.PRNG")
			
			' Random bytes encoded as hex or base64 (base 64 by default)
			
			random_token = PRNG.RandomBytes(keyLen,encoding)
			if lCase(encoding) = "hex" then random_token = lCase(random_token)
			
		set PRNG = nothing
		
	end function
	
	Dim start
		
	start = Timer()
		
	response.write "<p><b>Random 16 char string:</b> " & random_string(16) & "</p>"
	
	response.write "<p><b>Random 16 char sanitized string:</b> " & random_string(16) & "</p>"
	
	response.write "<p><b>Random 12 char password:</b> " & random_password(12) & "</p>"
	
	response.write "<p><b>Random 5 digit number:</b> " & random_number(10000,99999) & "</p>"
	
	response.write "<p><b>Random 128 bit hex token:</b> " & random_token(16,"hex")
	
	response.write "<p><b>Random 128 bit base64 token:</b> " & random_token(16,"b64")
	
	response.write "<p><b>Time to execute:</b> " & formatNumber(Timer()-start,6) & "s</p>"

%>