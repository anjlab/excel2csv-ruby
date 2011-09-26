require "excel2csv/version"
require "excel2csv/info"

require "csv"
require "tmpdir"

module Excel2CSV
  
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
    if info && Dir.exists?(info.working_folder)
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
      info.clean if block_given? && info
    end
  end

  module_function :convert

  def path_to_sheet(info, options = {})
    options.delete(:encoding) # all previews are in utf-8
    if options[:preview]
      collection = info.previews
    else
      collection = info.sheets
    end
    index = (idx = options[:sheet]) ? idx : 0
    collection[index][:path]
  end

  module_function :path_to_sheet

  def create_cvs_files(path, options)
    tmp_dir = Dir.mktmpdir
    working_folder = options[:working_folder] || tmp_dir
    limit = options[:rows]
    if path =~ /\.csv$/
      total_rows = 0
      preview_rows = []
      opts = clean_options(options)

      # Transcode file to utf-8, count total and gen preview

      CSV.open("#{working_folder}/1-#{total_rows}.csv", "wb") do |csv|
        CSV.foreach(path, opts) do |row| 
          if limit && total_rows <= limit
            preview_rows << row
          end
          total_rows += 1
          csv << row
        end
      end
      
      if limit
        CSV.open("#{working_folder}/1-#{limit}-of-#{total_rows}-preview.csv", "wb") do |csv|
          preview_rows.each {|row| csv << row}
        end
      end
    else
      java_options = options[:java_options] || "-Dfile.encoding=utf8 -Xms512m -Xmx512m -XX:MaxPermSize=256m"
      rows_limit = limit ? "-r #{limit}" : ""
      jar_path = File.join(File.dirname(__FILE__), "excel2csv.jar")
      `java #{java_options} -jar #{jar_path} #{rows_limit} #{path} #{working_folder}`
    end
    
    Info.read(working_folder)
  end

  module_function :create_cvs_files

  def clean_options options
    options.dup.delete_if do |key, value|
      [:working_folder, :java_options, :preview, :sheet, :path, :rows, :info].include?(key)
    end
  end

  module_function :clean_options

end
