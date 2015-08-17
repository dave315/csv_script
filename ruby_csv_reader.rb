require 'roo'
require 'csv'

file_to_open = 'csv_files/staples.xlsx'

xlsx = Roo::Spreadsheet.open(file_to_open)
xlsx.default_sheet = xlsx.sheets.last

new_array = Array.new
attribute_regexp = /art/

CSV.parse(xlsx.to_csv, headers:true) do |row|
  CSV.open('csv_files/output.csv', 'a+') do |csv|
    if row['Class_Name'] =~ attribute_regexp
      csv << row
    end
  end
end
