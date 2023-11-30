function Convert-PPTToImages {
    param (
        [string]$pptFile,
        [string]$outputFolder
    )
    $application = New-Object -ComObject powerpoint.application
    $pres = $application.Presentations.Open($pptFile)

    If(!(test-path -PathType container $outputFolder)){
      New-Item -ItemType Directory -Path $outputFolder
    }
    
    for ($i = 1; $i -le $pres.Slides.Count; $i++) {
        $slide = $pres.Slides.Item($i)
        $outputPath = Join-Path $outputFolder ("Slide" + $i + ".png")
        $slide.Export($outputPath, "PNG")
    }

    $pres.Close()
    $application.Quit()

    [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$pres) | out-null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

$pptFile = "$pwd\Lecture5 - Genetic Distance.pptx"
$outputFolder = "$pwd\images"
Convert-PPTToImages -pptFile $pptFile -outputFolder $outputFolder