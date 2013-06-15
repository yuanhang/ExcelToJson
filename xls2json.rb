require 'spreadsheet'

class Event
  #class << self
    #attr_accessor :attributes
  #end
  @@attributes = [:id, :content, :strategy, :product, :tech, :operation, :fortune, :range, :reusable, :comment]
  @@attributes.each { |attr| attr_accessor attr }
  def self.attributes
    @@attributes
  end
  def key attr
    instance_variable_get "@#{attr}"
  end
  #attr_accessor :id, :content, :strategy, :product, :tech, :operation, :fortune, :range, :reusable, :comment
  def initialize values  
    i = 0
    @id = values[i].to_i;
    i += 1
    @content = values[i]
    i += 1
    @strategy = values[i].to_i
    i += 1
    @product = values[i].to_i
    i += 1
    @tech = values[i].to_i
    i += 1
    @operation = values[i].to_i
    i += 1
    @fortune = values[i].to_i
    i += 1
    @range = values[i]
    i += 1
    @reusable = values[i]
    i += 1
    @comment = values[i]
  end

  def convert_to_point str
    if str.empty?
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
  rows << row;
  events << Event.new(row)
  #puts events.last.to_s
end

# Create output json file
json_file_name = xls_file_name.split('.')[0] + '.json'
File.delete(json_file_name) if File.exist?(json_file_name)

json_file = File.open(json_file_name, 'w')
json_file << "[\n"

#rows.each_with_index do |row, index1|
#json_file << "{\n"
#Event.attributes.each_with_index do |key, index|
#json_file << "#{key}:#{row[index]},\n"
##puts row[index]
#end
#json_file << "},\n"
#end
events.each_with_index do |event, index1|
  json_file << "{\n"
  Event.attributes.each_with_index do |key, index|
    json_file << "#{key}:#{event.key(key)},\n"
    #puts row[index]
  end
  json_file << "},\n"
end

json_file << "]\n"
json_file.close
puts events.to_s

