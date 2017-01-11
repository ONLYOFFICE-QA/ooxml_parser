module OoxmlParser
  # Class for storing Start data, `start` tag
  class Start < OOXMLDocumentObject
    # @return [Integer] value of start
    attr_accessor :value

    # Parse Start
    # @param [Nokogiri::XML:Node] node with Start
    # @return [Start] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'val'
          @value = value.value.to_f
        end
      end
      self
    end
  end
end
