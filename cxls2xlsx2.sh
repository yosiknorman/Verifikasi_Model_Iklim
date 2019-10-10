#!/bin/bash
IFS=$'\n'
mm=`date '+%m'`
mmin=`date '+%m' -d "2 month ago"`
# echo $mm

for i in `ls input/2019 | head -4l`
do    
    cd "input/2019/$i/Sifat"
    rm sifat\ hujan.2019.xls
    rm sifat\ hujan.new2019.xls
    libreoffice --headless --convert-to xls sifat\ hujan.2019.xlsx
    libreoffice --headless --convert-to xls sifat\ hujan.new2019.xlsx
    cd ../../../..  

    cd "input/2019/$i/Curah Hujan"
    rm curah\ hujan.2019.xls
    libreoffice --headless --convert-to xls curah\ hujan.2019.xlsx
    cd ../../../..

done

j=`ls input | grep "$mm. "` 
k=`ls input | grep "$mmin. "` 
cd input
echo $j
    # rm "$j"x
    # libreoffice --headless --convert-to xlsx $j
    # rm "$k"x
echo $k
    # libreoffice --headless --convert-to xlsx $k
cd ..

./verifikasi_RR.R
./verifikasi_RR_anal.R

