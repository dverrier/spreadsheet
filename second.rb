require 'rubygems'
require 'rubyXL'
require 'HtmlWriter'
require 'awesome_print'

def describe_workbook(workbook)
  workbook.worksheets.each {|sheet| puts "Sheet name #{sheet.sheet_name}"}
end

def load_big
  puts "Loading big Spreadsheet"
  source = RubyXL::Parser.parse 'voyage.xlsx'
  puts "there are #{source.worksheets.length} sheets"
  describe_workbook(source)
  source
end 

def load_little
  puts "Loading little Spreadsheet"
  source = RubyXL::Parser.parse 'little.xlsx'
  puts "there are #{source.worksheets.length} sheets"
  describe_workbook(source)
  source
end 


source = load_big
#source = load_little

#target = RubyXL::Workbook.new




@max_sheets = source.worksheets.length
0.upto(@max_sheets) do |sheet|
  #target.worksheets << RubyXL::Worksheet.new( 'mysheet'  )
  #target.worksheets[sheet].sheet_name = (sheet_name = 'mysheet' + sheet.to_s)
  source_sheet = source.worksheets[sheet]
  source_sheet_values = source_sheet.extract_data
  @max_row = (source_sheet_values.length) - 1
  @max_col = 9 #index of last ruby array member

  if source_sheet.sheet_data[4][0]
    @file_name = source_sheet.sheet_data[4][0].value
  else
    @file_name = "INDEX"
  end

  @h = HtmlWriter.new @file_name
  @h.start_html_file 
  0.upto(@max_row) do |row| 
    @h.start_table_row
    0.upto(@max_col) do |col|
      puts "sheet #{sheet} name #{source_sheet.sheet_name} rows #{@max_row} row #{row} col #{col}"
      @row = source_sheet.sheet_data[row]
      if @row
        @cell = source_sheet.sheet_data[row][col]
      end

      if @cell then
        if @cell.datatype=='s' then
          ap @cell.datatype
          @h.table_cell @cell.value
        else
          @h.table_cell ""
        end
      else
          @h.table_cell ""
      end
    end
    @h.end_table_row
  end
  @h.end_html_file
end


# puts "New spreadsheet there are #{target.worksheets.length} sheet(s)"
# describe_workbook(target)
# target.write '/home/dverrier/spreadsheet/voy1.xlsx'