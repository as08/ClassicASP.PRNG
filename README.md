This is a Component Object Model (COM) Dynamic-link library (DLL) coded in C# that can be set in Classic ASP using VBscripts "CreateObject" method and allows you to use the .NET RNGCryptoServiceProvider class to generate cryptographically secure pseudorandom numbers and strings. It should be used in place of VBscripts Rnd() function.

Uses the .NET RNGCryptoServiceProvider class
https://docs.microsoft.com/en-us/dotnet/api/system.security.cryptography.rngcryptoserviceprovider?view=netframework-4.7.2

## INSTALLATION

Make sure you have the lastest .NET Framework installed (tested on .NET Framework 4.7.2)

Open IIS, go to the applicaton pools and open the application pool being used by your Classic ASP application. Check the .NET CRL version E.g: v4.0.30319

Navigate to the CRL folder E.g: C:\Windows\Microsoft.NET\Framework64\v4.0.30319

Copy over: ClassicASP.PRNG.dll

Run CMD as administrator

Change the directory to your CRL folder E.g: cd C:\Windows\Microsoft.NET\Framework64\v4.0.30319

Run the following command: RegAsm ClassicASP.PRNG.dll /tlb /codebase

## Usage

	Set PRNG = Server.CreateObject("ClassicASP.PRNG")

	' RandomString([string length],[characters to use])
	' Default characters: abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890
	PRNG.RandomString(16) ' random 16 character string

	' RandomNumber([min value],[max value])
	PRNG.RandomNumber(1,10) ' random number between 1 and 10 (including 10)

	' RandomBytes([number of bytes],[encoding])
	' encoding can be hex or base64 (base64 by default)
	PRNG.RandomBytes(16,"hex") ' random 128 bit hex token
  
## Example output from PRNG.asp
  
	Random 16 char string: nJlhVdBNoqKBJ0Jk

	Random 16 char sanitized string: ok1ZbkREcHPMjXtR

	Random 12 char password: L*VlP&pKQ:-^

	Random 5 digit number: 84840

	Random 128 bit hex token: 9f766c312a56e151b9b704c3591160d0

	Random 128 bit base64 token: by/XZ9Zhnl+NHroUfCU2rA==

	Time to execute: 0.0001s
