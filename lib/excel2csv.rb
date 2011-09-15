require "excel2csv/version"
require "csv"
require "tmpdir"
require "fileutils"

module Excel2CSV

  class Info
    attr_accessor :sheets
    attr_accessor :tmp_dir
    attr_accessor :working_dir

    def self.read dir, tmp_dir
      info = Info.new dir, tmp_dir
      info.read
      info
    end

    def read
      @sheets = Dir["#{@working_dir}/*.csv"].map do |file|
        {path: file}
      end
    end 

    def close
      FileUtils.remove_entry_secure(@tmp_dir, true) if @tmp_dir
    end

    private

    def initialize working_dir, tmp_dir
      @working_dir = working_dir
      @tmp_dir = tmp_dir
      @sheets = []
    end

  end
  
  def foreach(path, options = {}, &block) 
    convert(path, options) do |info|
      CSV.foreach(info.sheets.first[:path], options, &block)
    end 
  end

  module_function :foreach

  def read(path, options = {})
    convert(path, options) do |info|
      CSV.read(info.sheets.first[:path], options)
    end
  end

  module_function :read

  def convert(path, options = {})
    begin
      tmp_dir = Dir.mktmpdir
      dest_folder = options[:dest_folder] || tmp_dir
      java_options = options[:java_options] || "-Dfile.encoding=utf8 -Xms512m -Xmx512m -XX:MaxPermSize=256m"
      `java #{java_options} -jar lib/excel2csv.jar #{path} #{dest_folder}`
      info = Info.read dest_folder, tmp_dir
      if block_given?
        yield info
      else
        info
      end
    ensure
      info.close if block_given?
    end
  end

  module_function :convert

end
