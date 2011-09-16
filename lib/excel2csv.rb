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
      CSV.foreach(path_to_sheet(info, options), clean_options(options), &block)
    end 
  end

  module_function :foreach

  def read(path, options = {})
    convert(path, options) do |info|
      CSV.read(path_to_sheet(info, options), clean_options(options))
    end
  end

  module_function :read

  def convert(path, options = {})
    info = options[:info]
    if info && Dir.exists?(info.working_dir)
      return block_given? ? yield(info) : info
    end
    begin
      info = create_cvs_files(path, options)
      if block_given?
        yield info
      else
        info
      end
    ensure
      info.close if block_given? && info
    end
  end

  module_function :convert

  def path_to_sheet(info, options = {})
    if options[:preview]
      options.delete(:encoding) # all previews are in utf-8
      collection = info.previews
    else
      collection = info.sheets
    end
    index = (idx = options[:index]) ? idx : 0
    collection[index][:path]
  end

  module_function :path_to_sheet

  def create_cvs_files(path, options)
    tmp_dir = Dir.mktmpdir
    dest_folder = options[:dest_folder] || tmp_dir
    java_options = options[:java_options] || "-Dfile.encoding=utf8 -Xms512m -Xmx512m -XX:MaxPermSize=256m"
    rows_limit = (limit = options[:rows_limit]) ? "-r #{limit}" : ""
    
    if path =~ /\.csv/
      total_rows = 0
      preview_rows = []
      opts = clean_options(options)
      CSV.foreach(path, opts) do |row| 
        if limit && total_rows <= limit
          preview_rows << row
        end
        total_rows += 1
      end
      FileUtils.cp path, "#{dest_folder}/1-#{total_rows}.csv"
      if limit
        CSV.open("#{dest_folder}/1-#{limit}-of-#{total_rows}-preview.csv", "wb") do |csv|
          preview_rows.each {|row| csv << row}
        end
      end
    else
      jar_path = File.join(File.dirname(__FILE__), "excel2csv.jar")
      `java #{java_options} -jar #{jar_path} #{rows_limit} #{path} #{dest_folder}`
    end
    
    Info.read(dest_folder, tmp_dir)
  end

  module_function :create_cvs_files

  def clean_options options
    options.dup.delete_if do |key, value|
      [:dest_folder, :java_options, :preview, :index, :path, :rows_limit, :info].include?(key)
    end
  end

  module_function :clean_options

end
