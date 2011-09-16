#encoding: utf-8

require 'excel2csv'

describe Excel2CSV do

  let(:excel) {Excel2CSV}

  it "reads xls files" do
    data = excel.read "spec/fixtures/basic_types.xls"
    data[0].should == ["1.00", "2011-12-23 21:00:00 UTC(+0000)", "Hello"]
    data[1].should == ["2.00", "2011-12-24 21:00:00 UTC(+0000)", "Привет"]
    data[2].should == ["3.00", "2011-12-25 21:00:00 UTC(+0000)", 'Привет, "я excel!"']
  end

  it "reads xlsx files" do
    data = excel.read "spec/fixtures/basic_types.xlsx"
    data[0].should == ["1.00", "2011-12-23 21:00:00 UTC(+0000)", "Hello"]
    data[1].should == ["2.00", "2011-12-24 21:00:00 UTC(+0000)", "Привет"]
    data[2].should == ["3.00", "2011-12-25 21:00:00 UTC(+0000)", 'Привет, "я excel!"']
  end

  it "iterates rows" do
    count = 0
    excel.foreach "spec/fixtures/basic_types.xls" do |row|
      row.length.should == 3
      count += 1
    end
    count.should == 3
  end

  it "removes tmp dir after work" do
    tmp_dir = nil
    excel.convert "spec/fixtures/basic_types.xlsx" do |info|
      # puts IO.read(info.sheets.first[:path])
      tmp_dir = info.tmp_dir
    end
    tmp_dir.should_not be_nil
    Dir.exists?(tmp_dir).should == false
  end

  it "converts once if info is passed" do
    info = excel.convert "spec/fixtures/basic_types.xlsx"
    info.sheets.length.should == 1
    info.previews.length.should == 0
    info.should == excel.convert("spec/fixtures/basic_types.xlsx", info:info)
  end

  it "regenerate csv files if working_dir is removed" do
    info = excel.convert "spec/fixtures/basic_types.xlsx"
    info.close
    info.should_not == excel.convert("spec/fixtures/basic_types.xlsx", info:info)
  end

  it "generates preview csv files with rows limit" do
    info = excel.convert "spec/fixtures/basic_types.xls", rows_limit:1
    info.sheets.length.should == 1
    info.previews.length.should == 1

    info.sheets.first[:total_rows].should == 3
    info.previews.first[:total_rows].should == 3

    info.sheets.first[:rows].should == 3
    info.previews.first[:rows].should == 1
  end

  it "reads previews" do
    data = excel.read("spec/fixtures/basic_types.xls", rows_limit:1, preview:true, index:0)
    data.length.should == 1
    data[0].should == ["1.00", "2011-12-23 21:00:00 UTC(+0000)", "Hello"]
  end

  it "reads csv files" do
    data = excel.read("spec/fixtures/basic_types.csv", encoding:'windows-1251:utf-8')
    data[0].should == ["1.00","12/24/11 12:00 AM","Hello"]
    data[1].should == ["2.00","12/25/11 12:00 AM","Привет"]
    data[2].should == ["3.00","12/26/11 12:00 AM",'Привет, "я excel!"']
  end

  it "reads csv files with preview" do
    data = excel.read("spec/fixtures/basic_types.csv", encoding:'windows-1251:utf-8',rows_limit:2,preview:true)
    data[0].should == ["1.00","12/24/11 12:00 AM","Hello"]
    data[1].should == ["2.00","12/25/11 12:00 AM","Привет"]
  end

  # Date, Boolean, String, [Phone, Percent, Email, Gender, Url]
  # Несколько телефонов, 


end