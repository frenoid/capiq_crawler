$ErrorActionPreference = "SilentlyContinue"
Stop-Transcript | out-null
$ErrorActionPreference = "Continue"
Start-Transcript -path output.txt -append

$mass_no = 10
$codes = Get-Content -Path gic_codes_to_download.txt

"Downloading " + $codes.length + " codes. Company Screening information" 
"$codes"

foreach ($code in $codes) {
		"***** Downloading $code Mass Variables"
		python mass_screening.py "$code" $mass_no all

}

"End of script"

Stop-Transcript


