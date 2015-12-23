require_relative 'line_end/line_size'
module OoxmlParser
  class LineEnd < OOXMLDocumentObject
    attr_accessor :type, :length, :width

    def self.parse(line_end_node)
      line_end = LineEnd.new
      line_end_node.attributes.each do |key, value|
        case key
        when 'type'
          line_end.type = value.value.to_sym
        when 'w'
          line_end.width = LineSize.parse(value.value)
        when 'len'
          line_end.length = LineSize.parse(value.value)
        end
      end
      line_end
    end
  end
end
