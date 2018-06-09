module OoxmlParser
  # Class for parsing `cond` tags
  class Condition < OOXMLDocumentObject
    attr_accessor :event, :delay, :duration

    # Parse Condition
    # @param node [Nokogiri::XML::Element] node to parse
    # @return [Condition] value of SheetFormatProperties
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'evt'
          @event = value.value
        when 'delay'
          @delay = value_to_symbol(value)
        end
      end
      self
    end
  end
end
