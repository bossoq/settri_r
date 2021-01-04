## load library
library(XML)
library(RCurl)
library(rlist)
library(dplyr)
library(jsonlite)

## set work dir
setwd(dir = "/workspaces/SETMKT") ## set ur current working dir

## function import data
import_xml <- function(url) {
    import_url <- getURL(url, .opts = list(ssl.verifypeer = FALSE))
    import_data <- readHTMLTable(import_url)
    import_data <- list.clean(import_data, fun = is.null, recursive = FALSE)
    import_data <- as.data.frame(import_data[2])
    import_data <- setNames(import_data, header)
    import_data <- replace(import_data, import_data == "", NA)
    import_data <- mutate_at(import_data, -1,
        function(x) as.numeric(gsub(",", "", x)))
}

## generate return function
get_return <- function(year, period, index) {
    ## retrieve rf data
    rf_url <- paste0("http://www.thaibma.or.th/yieldcurve/getintpttm?year=",
        year)
    rf_data <- fromJSON(content(GET(rf_url, accept_json()), "text"))
    rf_data["date"] <- as.Date(rf_data[, 1])
    rf_data[1] <- as.numeric(format(as.Date(rf_data[, 1]), "%Y%m%d"))
    rf_data <- subset(rf_data, !duplicated(substr(asof, 1, 6), fromLast = TRUE))
    rf_data[1] <- format(rf_data[, "date"], "%b-%Y")
    rf_data <- rf_data[, -55]

    ## loop through month
    for (i in month_c) {

        ## calc neccessary info
        asof <- paste0(i, "-", year)
        start_row <- which(ret_index_data[1] == asof)
        max_row <- colSums(!is.na(ret_index_data[index])) - start_row + 1
        end_year <- min(max_row, tail(period, n = 1) * 12) %/% 12
        period <- period[period <= end_year]

        ## calc ret index
        tmp <- c()
        for (j in period) {
            tmp <- append(tmp,
                mean(ret_index_data[start_row:((j * 12) +
                    start_row - 1), index]))
        }

        ## calc div yield
        tmp1 <- c()
        for (k in period) {
            tmp1 <- append(tmp1,
                mean(dividend_data[start_row:((k * 12) +
                    start_row - 1), index]))
        }

        ## calc tri
        tmp2 <- tmp + tmp1

        ## calc rf
        tmp3 <- c()
        for (l in period) {
            rfrow <- which(rf_data[1] == asof)
            tmp3 <- rf_data[rfrow, paste0(l, "Y")] / 100
        }
        result <- data.frame(period, tmp, tmp1, tmp2, tmp3)
        result <- setNames(result, c("Year",
            paste0(index, " market return"),
            paste0(index, " dividend yield"),
            paste0(index, " total market return"),
            "Risk free rate"))
        filename <- paste0(index, " Market return ", year, ".xlsx")
        write.xlsx(result, filename,
            sheetName = asof,
            col.names = TRUE, row.names = FALSE,
            append = TRUE)
    }
}

## set header
header <- c("Month-Year",
    "SET",
    "SET50",
    "SET100",
    "sSET",
    "SETCLMV",
    "SETHD",
    "SETTHSI",
    "SETWB",
    "mai")
month_c <- c("Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec")

## import data
dividend_url <- "https://www.set.or.th/static/mktstat/Table_Yield.xls"
dividend_data <- import_xml(dividend_url)
dividend_data[, -1] <- dividend_data[, -1, drop = FALSE] / 100
index_url <- "https://www.set.or.th/static/mktstat/Table_Index.xls"
index_data <- import_xml(index_url)
index_data <- index_data[-c(1, 2), ]

## create ret index
nrow <- colSums(!is.na(index_data))[1] - 12
ncol <- dim(index_data)[2]
ret_index_data <- index_data[1:nrow, 1, drop = FALSE]
for (i in 2:ncol) {
    nrow <- colSums(!is.na(index_data))[i]
    tmp_index_data <- index_data[1:(nrow - 12), 1, drop = FALSE]
    tmp_index_data[header[i]] <-
        (index_data[1:(nrow - 12), i] -
        index_data[13:nrow, i]) /
        index_data[13:nrow, i]
    ret_index_data <- left_join(ret_index_data,
        tmp_index_data,
        by = "Month-Year")
}

## query info
year <- 2020 ## input query year here
period <- c(1, 3, 5, 10, 17, 19, 20, 25, 30, 40, 42) ## input query period here
index <- "SET" ## input query market here

## run function to create xlsx
get_return(year, period, index)
