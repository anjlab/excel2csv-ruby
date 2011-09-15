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

end