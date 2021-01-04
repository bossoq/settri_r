# settri_r
Retrieve Thailand TRI and Gov Bond and export to excel by R

This code retrieve data directly from SET

Market Index : https://www.set.or.th/static/mktstat/Table_Index.xls

Dividend Yield : https://www.set.or.th/static/mktstat/Table_Yield.xls


and calculate TRI using yearly market return plus dividend yield :)

This code also retrieve government bond yield directly from ThaiBMA

Goverment Bond Yield : http://www.thaibma.or.th/EN/Market/YieldCurve/Government.aspx

to use it you need to set these variables:
1. setwd(dir = 'your current working dir')
2. year <- 'query year'
3. period <- c('period') ## yearly period
4. index <- 'your desired market'

Then, run this code and it will generate .XLSX file which contains TRI and Bond Yield for all months in that period.

Feel free to use this code :)
