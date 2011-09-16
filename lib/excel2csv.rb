require "excel2csv/version"
require "csv"
require "tmpdir"
require "fileutils"

module Excel2CSV

  class Info
    attr_accessor :sheets
    attr_accessor :previews
    attr_accessor :tmp_dir
    attr_accessor :working_dir

    def self.read dir, tmp_dir
      info = Info.new dir, tmp_dir
      info.read
      info
    end

    def read
      Dir["#{@working_dir}/*.csv"].map do |file|
        name = File.basename(file)
        m = /(?<sheet>\d+)-(?<rows>\d+)(-of-(?<total_rows>\d+))?/.match(name)
        next if !m
        total_rows   = (m[:total_rows] || m[:rows]).to_i
        preview_rows = m[:rows].to_i
        if name =~ /preview/
          @previews << {path: file, total_rows:total_rows, rows:preview_rows}
        else
          @sheets << {path: file, total_rows:total_rows, rows:total_rows}
        end
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
      @previews = []
    end

  end
  
  def foreach(path, options = {}, &block) 
    convert(path, options) do |info|
      CSV.foreach(path_to_sheet(info, options), options, &block)
    end 
  end

  module_function :foreach

  def read(path, options = {})
    convert(path, options) do |info|
      CSV.read(path_to_sheet(info, options), options)
    end
  end

  module_function :read

  def convert(path, options = {})
    info = options.delete(:info)
    if info && Dir.exists?(info.working_dir)
      return block_given? ? yield(info) : info
    end
    begin
      tmp_dir = Dir.mktmpdir
      jar_path = File.join(File.dirname(__FILE__), "excel2csv.jar")
      dest_folder = options.delete(:dest_folder) || tmp_dir
      java_options = options.delete(:java_options) || "-Dfile.encoding=utf8 -Xms512m -Xmx512m -XX:MaxPermSize=256m"
      rows_limit = (limit = options.delete(:rows_limit)) ? "-r #{limit}" : ""
      `java #{java_options} -jar #{jar_path} #{rows_limit} #{path} #{dest_folder}`
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

  def path_to_sheet(info, options)
    collection = options.delete(:preview) ? info.previews : info.sheets
    index = (idx = options.delete(:index)) ? idx : 0
    collection[index][:path]
  end

  module_function :path_to_sheet

end
