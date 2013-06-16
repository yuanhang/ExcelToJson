require 'spreadsheet'

class Event
  #class << self
  #attr_accessor :attributes
  #end
  @@attributes = [:id, :title, :content, :strategy, :product, :tech, :operation, :fortune, :range, :reusable, :comment]
  @@attributes.each { |attr| attr_accessor attr }
  def self.attributes
    @@attributes
  end
  def value_for_key key
    instance_variable_get "@#{key}"
  end
  def initialize values  
    i = 0
    @id = values[i].to_i;
    i += 1
    if (values[i] == nil)
      puts @id
    end
    @title = values[i].split("\n")[0]
    @content = values[i].split("\n")[1]
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

  def to_json
    json_str = ""
    @@attributes.each_with_index do |key, i|
      json_str += "\"#{key}\":\"#{value_for_key(key)}\""
      json_str += "," unless i == @@attributes.size-1
    end
    json_str
  end

end

xls_file_name = 'events.xls'
Spreadsheet.client_encoding = 'UTF-8'
book = Spreadsheet.open(xls_file_name)
sheet1 = book.worksheet(0)

omit_rows = 2
keys_row_index = 1

events = Array.new
sheet1.each omit_rows do |row|
  events << Event.new(row)
end

# Create output json file
json_file_name = xls_file_name.split('.')[0] + '.json'
File.delete(json_file_name) if File.exist?(json_file_name)

json_file = File.open(json_file_name, 'w')
json_file << "{\"events\":["
events.each_with_index do |event, i|
  json_str = "{#{event.to_json}}" 
  json_str += "," unless i == events.size-1
  json_file << json_str
end
json_file << "]}"
json_file.close
puts events[0].to_json

