require 'spreadsheet'

class Event
  def initialize values  
    i = 0
    @id = Integer(values[i]);
    ++i
    @content = values[i]
    ++i
    @strategy = convert_to_point values[i]
    ++i
    @product = convert_to_point values[i]
    ++i
    @tech = convert_to_point values[i]
    ++i
    @operation = convert_to_point values[i]
    ++i
    @fortune = convert_to_point values[i]
    ++i
    @range = values[i]
    ++i
    @reusable = values[i]
    ++i
    @comment = values[i]
  end

  def convert_to_point str
    if String.empty?
      return 0
    else
      return str.to_i
    end
  end
end

xls_file_name = 'events.xls'
Spreadsheet.client_encoding = 'UTF-8'
book = Spreadsheet.open(xls_file_name)
sheet1 = book.worksheet(0)

omit_rows = 2
keys_row_index = 1
keys = Array.new
keys = sheet1.row(keys_row_index)
puts keys
rows = Array.new
events = Array.new
sheet1.each omit_rows do |row|
  rows << row
  events << Event.new row
end

# Create output json file
json_file_name = xls_file_name.split('.')[0] + '.json'
File.delete(json_file_name) if File.exist?(json_file_name)

json_file = File.open(json_file_name, 'w')
json_file << "[\n"

rows.each_with_index do |row, index|
  json_file << "{\n"
  keys.each_with_index do |key, index|
    json_file << "#{key}:#{row[index]},\n"
    puts row[index]
  end
  json_file << "},\n"
end

json_file << "]\n"
json_file.close
puts events

