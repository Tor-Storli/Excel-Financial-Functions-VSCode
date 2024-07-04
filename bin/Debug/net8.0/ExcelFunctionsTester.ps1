
$path = "C:\Users\storl\Financial\ExcelFunctions\bin\Debug\net8.0"
Start-Process -FilePath "$path\ExcelFunctions.exe" -ArgumentList "--excelfunction NPer --periodicInterestRate (.0325/12) --pmt_value -2459.34 --pv_value 350000 --fv_value 0 --beg_end 0" -Verb RunAs
