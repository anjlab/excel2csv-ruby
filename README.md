# Excel2CSV: gem for converting Excel files to csv

### Installation and usage

``` ruby
gem install excel2csv
```

With Gemfile

``` ruby
gem 'excel2csv'
```

### Requirements

1. Ruby > 1.9.2
2. Java Runtime

### Usage

``` ruby
require 'excel2csv'

# Read csv, xls, xlsx files like with CSV#read
Excel2CSV.read "path/to/file.csv", encoding:"windows-1251:utf-8"
# by default Excel2CSV reads first worksheet
Excel2CSV.read "path/to/file.xls"  # working encoding is always UTF-8
Excel2CSV.read "path/to/file.xlsx" # working encoding is always UTF-8

# Line by line reading like with CSV#foreach
Excel2CSV.foreach("path/to/file.csv", encoding:"windows-1251:utf8") {|r| puts r}
Excel2CSV.foreach("path/to/file.xls") {|r| puts r}
Excel2CSV.foreach("path/to/file.xlsx") {|r| puts r}

# Read non first worksheet
Excel2CSV.read "path/to/file.xls", index:1 #reads second sheet 
Excel2CSV.read "path/to/file.xlsx", index:2 #reads third sheet


# Preview first N rows in big files
Excel2CSV.read "path/to/file.xls", rows_limit: 2, preview: true
Excel2CSV.read "path/to/file.xlsx", rows_limit: 2, preview: true
Excel2CSV.read "path/to/file.csv", rows_limit: 2, preview: true
```