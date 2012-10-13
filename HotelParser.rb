require 'rubygems'
require 'rubyXL'
require 'HtmlWriter'
require 'awesome_print'

class HotelParser

  def initialize size
    load_file size
    export
  end

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

  def load_file( size )
    if size == :big 
      @source = load_big
    else
      @source = load_little
    end
  end
#target = RubyXL::Workbook.new

  def get_file_name(sheet_number)
      case
      when sheet_number == 0
        name =  "INDEX"
      when sheet_number == 311
          name = "Visas"
      when sheet_number == 312
          name = "Transfers"
      when @source_sheet.sheet_data[4][0]
        name =  @source_sheet.sheet_data[4][0].value #'sheet'+ sheet.to_s
      else
        name = 'sheet'+ sheet_number.to_s
      end
  end

  def export_cell(row, col)
    #puts "sheet #{sheet} name #{@source_sheet.sheet_name} rows #{@max_row} row #{row} col #{col}"
    @row = @source_sheet.sheet_data[row]
    if @row
      @cell = @source_sheet.sheet_data[row][col]
    end
    if @cell then
      if @cell.datatype=='s' then
        #ap @cell.datatype
        @h.table_cell @cell.value
      else
        @h.table_cell ""
      end
    else
      @h.table_cell ""
     end
  end


  def export
    @max_sheets = @source.worksheets.length - 1 #index of last ruby array member - starting from zero
    0.upto(@max_sheets) do |sheet|
      @source_sheet = @source.worksheets[sheet]
      @source_sheet_values = @source_sheet.extract_data
      @max_row = (@source_sheet_values.length) - 1 #index of last ruby array member - starting from zero
      @max_col = 20 # was 9 #index of last ruby array member - starting from zero

      @file_name = get_file_name( sheet )

      puts "sheet #{sheet} name #{@source_sheet.sheet_name} rows #{@max_row}"
      @h = HtmlWriter.new @file_name
      @h.start_html_file 
      0.upto(@max_row) do |row| 
        @h.start_table_row
        0.upto(@max_col) do |col|
          export_cell(row, col)  

        end
        @h.end_table_row
      end
      @h.end_html_file
    end
  end #export

end #class



HotelParser.new :big
