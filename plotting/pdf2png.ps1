$fpath = $args[0]
$fname = Split-Path -Path $fpath -Leaf
$fname = $fname -replace "\..+"
$ffolder = Split-Path -Path $fpath
$cropped_fname = $fname + "-crop.pdf"
$cropped_fpath = Join-Path $ffolder $cropped_fname
pdfcrop $fpath $cropped_fpath
pdftoppm -r 300 -png $cropped_fpath (Join-Path $ffolder $fname)
rm $cropped_fpath