require 'fileutils'
require 'json'

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
      info = JSON.parse(File.read("#{@working_folder}/info.json"))
      info['sheets'].each do |sheet_info|
        @sheets << {
          path:       "#{@working_folder}/#{sheet_info['fullOutput']}",
          total_rows: sheet_info['rowCount'],
          rows:       sheet_info['rowCount'],
          title:      sheet_info['title']
        }
        
        @previews << {
          path:       "#{@working_folder}/#{sheet_info['previewOutput']}",
          total_rows: sheet_info['rowCount'],
          rows:       info['perSheetRowLimitForPreviews'],
          title:      sheet_info['title']
        } if sheet_info['previewOutput']
      end
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