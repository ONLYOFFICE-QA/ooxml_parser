module OoxmlParser
  # Class for parsing `lin` tags
  class LinearGradient < OOXMLDocumentObject
    attr_accessor :angle, :scaled

    # Parse LinearGradient object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [LinearGradient] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'ang'
          @angle = value.value.to_f / 100_000.0
        when 'scaled'
          @scaled = value.value.to_i
        end
      end
      self
    end
  end
end
