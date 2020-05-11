# ConvertTo-Jpeg - Converts RAW (and other) image files to the widely-supported JPEG format
# https://github.com/DavidAnson/ConvertTo-Jpeg

Param (
    [Parameter(
        Position = 1,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true,
        ValueFromRemainingArguments = $true,
        HelpMessage = "Array of image file names to convert to HEIC")]
    [Alias("FullName")]
    [String[]]
    $Files,

    [Parameter(
        HelpMessage = "Fix extension of HEIC files")]
    [Switch]
    $FixExtensionIfHeic,
    [Switch]
    $RemoveJpegAfterConverting
)
    #$Files=Get-ChildItem -Path $Folder\* -Include *.jpg, *.jpeg   -Recurse 
Begin
{
    # Technique for await-ing WinRT APIs: https://fleexlab.blogspot.com/2018/02/using-winrts-iasyncoperation-in.html
    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    $runtimeMethods = [System.WindowsRuntimeSystemExtensions].GetMethods()
    $asTaskGeneric = ($runtimeMethods | ? { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
    Function AwaitOperation ($WinRtTask, $ResultType)
    {
        $asTaskSpecific = $asTaskGeneric.MakeGenericMethod($ResultType)
        $netTask = $asTaskSpecific.Invoke($null, @($WinRtTask))
        $netTask.Wait() | Out-Null
        $netTask.Result
    }
    $asTask = ($runtimeMethods | ? { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' })[0]
    Function AwaitAction ($WinRtTask)
    {
        $netTask = $asTask.Invoke($null, @($WinRtTask))
        $netTask.Wait() | Out-Null
    }


    # Reference WinRT assemblies
    [Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime] | Out-Null
    [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics, ContentType=WindowsRuntime] | Out-Null
}

Process
{
    
    # Summary of imaging APIs: https://docs.microsoft.com/en-us/windows/uwp/audio-video-camera/imaging
    
    foreach ($file in $Files)
    {
        Write-Host $file -NoNewline               
        try
        {         
            try
            {       
                $ExifTool=(gci -Path exiftool.exe).FullName     
                $command = "$($ExifTool) -all:all= $('"'+$file+'"') --exif:Orientation -charset filename=cp1251"            
                $bytes = [System.Text.Encoding]::Unicode.GetBytes($command)
                $encodedCommand = [Convert]::ToBase64String($bytes)
                
                powershell -EncodedCommand $encodedCommand
                # Get SoftwareBitmap from input file
                $file = Resolve-Path -LiteralPath $file
                
                $inputfile = awaitoperation ([windows.storage.storagefile]::getfilefrompathAsync($file)) ([Windows.Storage.StorageFile])            
                $inputFolder = AwaitOperation ($inputFile.GetParentAsync()) ([Windows.Storage.StorageFolder])
                $inputStream = AwaitOperation ($inputFile.OpenReadAsync()) ([Windows.Storage.Streams.IRandomAccessStreamWithContentType])
                $decoder = AwaitOperation ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($inputStream)) ([Windows.Graphics.Imaging.BitmapDecoder])

                            
            }
            catch
            {
                # Ignore non-image files
                Write-Host " [Unsupported]"
                continue
            }
            if ($decoder.DecoderInformation.CodecId -eq [Windows.Graphics.Imaging.BitmapDecoder]::HeifDecoderId)
            {
                $extension = $inputFile.FileType                
                if ($FixExtensionIfHeic -and ($extension -ne ".heic") -and ($extension -ne ".heif"))
                {
                    # Rename HEIF-encoded files to have ".heic" extension
                    $newName = $inputFile.Name -replace ($extension + "$"), ".heic"
                    AwaitAction ($inputFile.RenameAsync($newName))
                    Write-Host " => $newName"
                }
                else
                {
                    # Skip JPEG-encoded files
                    Write-Host " [Already HEIC]"
                }
                continue
            }
            else{
                $extension = $inputFile.FileType                
                $outputFileName = $inputFile.Name.Replace($extension,".heic")
            }
            $bitmap = AwaitOperation ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])

            # Write SoftwareBitmap to output file
            #$outputFileName = $inputFile.Name + ".heic";
            $outputFile = AwaitOperation ($inputFolder.CreateFileAsync($outputFileName, [Windows.Storage.CreationCollisionOption]::ReplaceExisting)) ([Windows.Storage.StorageFile])
            $properOut = AwaitOperation ($outputFile.Properties.GetImagePropertiesAsync())([Windows.Storage.FileProperties.ImageProperties])                

            $outputStream = AwaitOperation ($outputFile.OpenAsync([Windows.Storage.FileAccessMode]::ReadWrite)) ([Windows.Storage.Streams.IRandomAccessStream])
            $encoder = AwaitOperation ([Windows.Graphics.Imaging.BitmapEncoder]::CreateAsync([Windows.Graphics.Imaging.BitmapEncoder]::JpegEncoderId, $outputStream)) ([Windows.Graphics.Imaging.BitmapEncoder])
            $encoder.SetSoftwareBitmap($bitmap)


            # Do it
            AwaitAction ($encoder.FlushAsync())
            Write-Host " -> $outputFileName"
        }
        catch
        {
            # Report full details
            throw $_.Exception.ToString()
        }
        finally
        {
            # Clean-up
            if ($inputStream -ne $null) { [System.IDisposable]$inputStream.Dispose() }
            if ($outputStream -ne $null) { [System.IDisposable]$outputStream.Dispose() }

            $properOut.Orientation.value__=6                      
        }
        try{
        .\metacopy -e -p $($file.Path+"_original") $outputFile.Path
        }
        catch{}
        if($RemoveJpegAfterConverting){
            Remove-Item $file.Path
            Remove-Item $($file.Path+"_original")
        }
        else{
            Remove-Item $file.Path
            Rename-Item  -Path $($file.Path+"_original") -NewName $($file.Path+"_original").Replace("_original","")
        }
    }
}