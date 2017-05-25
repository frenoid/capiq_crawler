$no = $args[0]
$download_dir = "D:\Capital IQ\info_extraction\mass_$no\raw_data"

python find_missing.py "$download_dir" screening
