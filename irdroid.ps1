param(
    [Parameter(Mandatory=$true)]
    [string[]]$Command
)

Set-Location $PSScriptRoot

# ---------------- ALIAS MAP ----------------
$CommandMap = @{
    "power"       	= "on_off"
    "onoff"       	= "on_off"
    "toggle"      	= "on_off"

    "hdmi1"         = "hdmi_1"
    "hdmi2"         = "hdmi_2"
    "hdmi3"         = "hdmi_3"
    "optical"       = "optical"
    "hdmitv"        = "hdmi_tv"
    "auxin"         = "aux_in"

    "bt"          	= "bluetooth"
    "bluetooth"   	= "bluetooth"

    "voldown"     	= "vol_down"
    "volumedown"  	= "vol_down"
    "quieter"     	= "vol_down"

    "volup"       	= "vol_up"
    "volumeup"    	= "vol_up"
    "louder"      	= "vol_up"

    "mute"        	= "de_mute"
    "unmute"      	=	"de_mute"

    "bassdown"     	= "bass_down"	

    "bassup"      	= "bass_up"

    "harmankardon"  = "harman_kardon"

    "stereo"      	= "stereo"

    "virtual"      	= "virtual"
	
    "wave"      	= "wave"

}

try {
    $pipe = New-Object IO.Pipes.NamedPipeClientStream(".", "irdroid", "InOut")
    $pipe.Connect(5000)

    $reader = New-Object IO.StreamReader($pipe)
    $writer = New-Object IO.StreamWriter($pipe)
    $writer.AutoFlush = $true

    $msgId = 0

    function Resolve-Command {
        param([string]$cmd)

        $key = $cmd.ToLower()

        if ($CommandMap.ContainsKey($key)) {
            return $CommandMap[$key]
        }

        # fallback: direkt verwenden (falls schon korrekt)
        return $cmd
    }

    function Send-Command {
        param([string]$cmd)

        $script:msgId++

        $request = @{
            id  = $script:msgId
            cmd = $cmd
        } | ConvertTo-Json -Compress

        $writer.WriteLine($request)

        $timeout = 5
        $sw = [Diagnostics.Stopwatch]::StartNew()

        while ($sw.Elapsed.TotalSeconds -lt $timeout) {
            try {
                if ($pipe.IsConnected -and $pipe.CanRead) {

                    $line = $reader.ReadLine()

                    if ($line) {
                        try {
                            $response = $line | ConvertFrom-Json
                            if ($response.id -eq $script:msgId) {
                                return $response.status
                            }
                        } catch {}
                    }
                }
            } catch {}

            Start-Sleep -Milliseconds 15
        }

        return "TIMEOUT"
    }

    foreach ($cmd in $Command) {

        $resolved = Resolve-Command $cmd

        if (-not $resolved) {
            Write-Host "Unknown command: $cmd"
            continue
        }

        $result = Send-Command $resolved
        Write-Host "$cmd ($resolved) : $result"

        Start-Sleep -Milliseconds 150
    }
}
catch {
    Write-Error "Pipe Fehler: $_"
}
finally {
    if ($reader) { $reader.Close() }
    if ($writer) { $writer.Close() }
    if ($pipe)   { $pipe.Close() }
}