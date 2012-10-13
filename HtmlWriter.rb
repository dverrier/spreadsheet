class HtmlWriter
  def initialize( name )
    @name = name
    @file = File.new( "html/" + name + ".html", "w")
    
  end

  def start_html_file
    @file.puts "<!doctype html>"
    @file.puts"<head>"
    @file.puts '<meta charset="utf-8">'
    @file.puts '<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">'
    @file.puts "<title> #{@name} </title>"
    @file.puts '<meta name="description" content="Latvia Tours UAE">'
    @file.puts '<link rel="stylesheet" href="css/style.css">'
    @file.puts '</head>'
    @file.puts '<body>'
    @file.puts '<table>'
  end

  def end_html_file
    @file.puts "</table>"
    @file.puts "</body>"
    @file.puts "</html>"
    @file.close
  end

  def start_table_row
  @file.puts "<tr>"
  end
  def end_table_row
    @file.puts "</tr>"
  end

  def table_cell content
    @file.puts "<td>"
    @file.puts content
    @file.puts "</td>"
  end

end