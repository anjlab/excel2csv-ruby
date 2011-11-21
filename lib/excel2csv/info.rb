require "fileutils"

module Excel2CSV
  class Info
    attr_accessor :sheets
    attr_accessor :previews
    attr_accessor :working_folder

    def self.read working_folder
      info = Info.new working_folder
      info.read
      info
    end

    def read
      Dir["#{@working_folder}/*.csv"].map do |file|
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
      @previews.sort! {|a, b| a[:path] <=> b[:path]}
      @sheets.sort! {|a, b| a[:path] <=> b[:path]}
    end 

    def clean
      FileUtils.remove_entry_secure(@working_folder, true) if @working_folder
    end

    private

    def initialize working_folder
      @working_folder = working_folder
      @sheets = []
      @previews = []
    end
  end
end