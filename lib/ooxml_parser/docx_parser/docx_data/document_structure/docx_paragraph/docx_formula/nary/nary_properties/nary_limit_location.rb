module OoxmlParser
  # Class for parsing `m:limLoc` object
  class NaryLimitLocation < OOXMLDocumentObject
    # @return [String] value of limit location
    attr_accessor :value

    # Parse NaryLimitLocation
    # @param [Nokogiri::XML:Node] node with NaryLimitLocation
    # @return [NaryLimitLocation] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'val'
          @value = value_to_symbol(value)
        end
      end
      self
    end
  end
end
