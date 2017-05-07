#シリアルポート
#起動パラメータ
#Param()

#GLOBAL Variables
[Object]$SerialPort = $null

#Types

Add-Type -TypeDefinition @"
   public enum RecievePhaseCode
   {
      PHASE_INITIALIZE,
      PHASE_READING,
      PHASE_REPLY,
      PHASE_MAX
   }
"@

#Functions
#ポート番号選択
Function SelectPort()
{
	$SerialPortNames = [System.IO.Ports.SerialPort]::getportnames()
	Write-Host ("Please Select Port. 0 - " + ([System.Array]$SerialPortNames).Count )
	for( $i = 0; $i -lt ([System.Array]$SerialPortNames).Count; $i++ )
	{
		Write-Host (([string]$i).PadRight(3) + ": " + [System.Array]$SerialPortNames[$i])
	}
	return ($SerialPortNames[(Read-Host "? >")])
}

Function Initialize( [ref]$SerialPort, [string]$PortName )
{
	if( $PortName -eq "" )
	{
		return $false
	}
	$i  = 0;
	$PortName | 
	% {
		[System.Array]$SerialPort.value += @{
			ID = $i;
			Obj = (new-Object System.IO.Ports.SerialPort $_,19200,None,8,one);
			Name = $_;
			ReadTimeout = 500;
			WriteTimeout = 100;
		}
		$SerialPort.value[$i].Obj.ReadTimeout = $SerialPort.value[$i].ReadTimeout;
		$SerialPort.value[$i].Obj.WriteTimeout = $SerialPort.value[$i].WriteTimeout;
		$SerialPort.value[$i].Obj.Open()
		$i++;
	}
	return $true
}



Function Finalize( [ref]$SerialPort )
{
	foreach( $Port in $SerialPort.value )
	{
		$Port.Obj.Close()
	}
}


Function Read-HostTimeout
{
###################################################################
##  Description:  Mimics the built-in "read-host" cmdlet but adds an expiration timer for
##  receiving the input.  Does not support -assecurestring
##
##  This script is provided as is and may be freely used and distributed so long as proper
##  credit is maintained.
##
##  Written by: thegeek@thecuriousgeek.org
##  Date Modified:  10-24-14
###################################################################

# Set parameters.  Keeping the prompt mandatory
# just like the original
param(
	[Parameter(Mandatory=$true,Position=1)]
	[string]$prompt,
	
	[Parameter(Mandatory=$false,Position=2)]
	[int]$delayInSeconds
)
	
	# Do the math to convert the delay given into milliseconds
	# and divide by the sleep value so that the correct delay
	# timer value can be set
	$sleep = 250
	$delay = ($delayInSeconds*1000)/$sleep
	$count = 0
	$charArray = New-Object System.Collections.ArrayList
	Write-host -nonewline "$($prompt):  "
	
	# While loop waits for the first key to be pressed for input and
	# then exits.  If the timer expires it returns null
	While ( (!$host.ui.rawui.KeyAvailable) -and ($count -lt $delay) ){
		start-sleep -m $sleep
		$count++
		If ($count -eq $delay) { "`n"; return $null}
	}
	
	# Retrieve the key pressed, add it to the char array that is storing
	# all keys pressed and then write it to the same line as the prompt
	$key = $host.ui.rawui.readkey("NoEcho,IncludeKeyUp").Character
	$charArray.Add($key) | out-null
	Write-host -nonewline $key
	
	# This block is where the script keeps reading for a key.  Every time
	# a key is pressed, it checks if it's a carriage return.  If so, it exits the
	# loop and returns the string.  If not it stores the key pressed and
	# then checks if it's a backspace and does the necessary cursor 
	# moving and blanking out of the backspaced character, then resumes 
	# writing. 
	$key = $host.ui.rawui.readkey("NoEcho,IncludeKeyUp")
	While ($key.virtualKeyCode -ne 13) {
		If ($key.virtualKeycode -eq 8) {
			$charArray.Add($key.Character) | out-null
			Write-host -nonewline $key.Character
			$cursor = $host.ui.rawui.get_cursorPosition()
			write-host -nonewline " "
			$host.ui.rawui.set_cursorPosition($cursor)
			$key = $host.ui.rawui.readkey("NoEcho,IncludeKeyUp")
		}
		Else {
			$charArray.Add($key.Character) | out-null
			Write-host -nonewline $key.Character
			$key = $host.ui.rawui.readkey("NoEcho,IncludeKeyUp")
		}
	}
	""
	$finalString = -join $charArray
	return $finalString
}

if( (Initialize -SerialPort ([ref]$SerialPort) -PortName (SelectPort)) -eq $false )
{
	Write-Host "初期化に失敗しました。"
	return
}

$CommandList = Import-CSV (( $profile | Split-Path ) + "\CameraCommandList.csv" )

foreach( $Port in [System.Array]$SerialPort )
{
	while( 1 )
	{
		Write-Host "受信状態のセットアップ中"
		$RecievePhase = [RecievePhaseCode]::PHASE_INITIALIZE
		$FirstByteRecievedFlag = $false
		[System.Array]$ReciveData = $null

		Write-Host "データ受信待機中"
		while( 1 )
		{
			if( $Port.Obj.BytesToRead -gt 0 )
			{
				Write-Host "データ受信開始"
				break
			}
		}

		while ( 1 )
		{
			try {
				#$Port.Obj.ReadChar()
				[System.Array]$ReciveData += $Port.Obj.ReadByte()
				$FirstByteRecievedFlag = $true
			} catch [exception] {
				if( $FirstByteRecievedFlag -eq $true )
				{
					$RecieveString = ""
					foreach( $Data in [System.Array]$ReciveData )
					{
						Write-Host ($Data.toString("X2")) -noNewLine
						Write-Host " " -noNewLine
						$RecieveString = $RecieveString + $Data.toString("X2") + " "
					}
					Write-Host ""
					Write-Host "データ受信終了"
					break
				}
			}
		}

		Write-Host "コマンド検出中"
		[Object]$SelectCommand = $null	
		if( $RecieveString -match "80 80 (.+) 81 81" )
		{
			$RecieveCommand = $matches[1]
			foreach( $Command in $CommandList )
			{
				if( $RecieveCommand -match $Command.RecieveCommand )
				{
					$SelectCommand = $Command
					break;
				}
			}
		}

		if( $SelectCommand -eq $null )
		{
			Write-Host "該当するコマンドはありませんでした。"
			continue
		}

		Write-Host "コマンドが見つかりました。"

		Write-Host $SelectCommand.Name

		Write-Host "応答メッセージの送信中"
		$SelectCommand.ReplyCommand -split " " |
		% {
			Write-Host ($_) -noNewLine
			Write-Host " " -noNewLine
			$Port.Obj.Write([Byte]([System.Convert]::ToInt32($_, 16)),0,1)
		}
		Write-Host ""
		Write-Host "応答メッセージの送信完了"
	}
}

Finalize ([ref]$SerialPort)
