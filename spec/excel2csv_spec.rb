#encoding: utf-8

require 'excel2csv'

describe Excel2CSV do

  let(:excel) {Excel2CSV}
  let(:plain)            {"spec/fixtures/plain.txt"}
  let(:csv_basic_types)  {"spec/fixtures/basic_types.csv"}
  let(:xls_basic_types)  {"spec/fixtures/basic_types.xls"}
  let(:xlsx_basic_types) {"spec/fixtures/basic_types.xlsx"}
  let(:csv_with_spaces)  {"spec/fixtures/file name with spaces.csv"}
  let(:xls_with_spaces)  {"spec/fixtures/file name with spaces.xls"}
  let(:xlsx_with_spaces) {"spec/fixtures/file name with spaces.xlsx"}
  let(:date_24) {Time.new(2011,12,24).strftime("%Y-%m-%dT%H:%M:%S")}
  let(:date_25) {Time.new(2011,12,25).strftime("%Y-%m-%dT%H:%M:%S")}
  let(:date_26) {Time.new(2011,12,26).strftime("%Y-%m-%dT%H:%M:%S")}

  it "reads plain txt files" do
    data = excel.read plain
    data[0].should == ["123"]
    data[1].should == ["456"]
  end

  it "reads xls files" do
    data = excel.read xls_basic_types
    data[0].should == ["1.00", date_24, "Hello"]
    data[1].should == ["2.00", date_25, "Привет"]
    data[2].should == ["3.00", date_26, 'Привет, "я excel!"']
  end

  it "reads xlsx files" do
    data = excel.read xlsx_basic_types
    data[0].should == ["1.00", date_24, "Hello"]
    data[1].should == ["2.00", date_25, "Привет"]
    data[2].should == ["3.00", date_26, 'Привет, "я excel!"']
  end

  it "iterates rows like CSV lib" do
    count = 0
    excel.foreach xls_basic_types do |row|
      row.length.should == 3
      count += 1
    end
    count.should == 3
  end

  it "removes tmp folder after work" do
    working_folder = nil
    excel.convert xlsx_basic_types do |info|
      # puts IO.read(info.sheets.first[:path])
      working_folder = info.working_folder
    end
    working_folder.should_not be_nil
    Dir.exists?(working_folder).should == false
  end

  it "converts once if info is passed" do
    info = excel.convert xlsx_basic_types
    info.sheets.length.should == 1
    info.previews.length.should == 0
    info.should == excel.convert(xlsx_basic_types, info:info)
  end

  it "regenerate csv files if working_folder is removed" do
    info = excel.convert xlsx_basic_types
    info.clean
    info.should_not == excel.convert(xlsx_basic_types, info:info)
  end

  it "generates preview csv files with rows limit" do
    info = excel.convert xls_basic_types, rows:1
    info.sheets.length.should == 1
    info.previews.length.should == 1

    info.sheets.first[:total_rows].should == 3
    info.previews.first[:total_rows].should == 3

    info.sheets.first[:rows].should == 3
    info.previews.first[:rows].should == 1
  end

  it "reads previews" do
    data = excel.read(xls_basic_types, rows:1, preview:true, sheet:0)
    data.length.should == 1
    data[0].should == ["1.00", date_24, "Hello"]
  end

  it "reads csv files" do
    data = excel.read(csv_basic_types, encoding:'windows-1251:utf-8')
    data[0].should == ["1.00","12/24/11 12:00 AM","Hello"]
    data[1].should == ["2.00","12/25/11 12:00 AM","Привет"]
    data[2].should == ["3.00","12/26/11 12:00 AM",'Привет, "я excel!"']
  end

  it "reads csv files with preview" do
    data = excel.read(csv_basic_types,
      encoding: 'windows-1251:utf-8',
      rows:     2,
      preview:  true
    )
    data[0].should == ["1.00","12/24/11 12:00 AM","Hello"]
    data[1].should == ["2.00","12/25/11 12:00 AM","Привет"]
  end

  context "file names with spaces" do
    it "reads csv files" do
      data = excel.read(csv_with_spaces, encoding:'windows-1251:utf-8')
      data.should == [["1.00"], ["2.00"]]
    end

    it "reads xls files" do
      data = excel.read(xls_with_spaces)
      data.should == [["1"], ["2"]]
    end

    it "reads xlsx files" do
      data = excel.read(xlsx_with_spaces)
      data.should == [["1"], ["2"]]
    end
  end


end