#require 'rubygems'

require 'spreadsheet'
require 'roo'
require 'awesome_print'

puts "Loading big Spreadsheet"
source = Excelx.new 'voyage.xlsx'
puts "there are #{source.sheets.length} sheets"


Spreadsheet.client_encoding = 'UTF-8'

target = Spreadsheet::Workbook.new

new_sheet = target.create_worksheet :name => source.sheets[1]
source.default_sheet = source.sheets[1]

1.upto(1000) do |row| 
  1.upto(100) do |col|
   new_sheet[row-1,col-1] = source.cell( row,col )

ap source.excelx_format( row,col )
ap source.cell( row,col )
exit

  end
end

target.write 'voy1.xls'
puts "New spreadsheet there are #{target.worksheets.length} sheet(s)"