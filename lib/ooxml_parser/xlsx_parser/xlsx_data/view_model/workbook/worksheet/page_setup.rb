module OoxmlParser
  # Class for parsing `pageSetup` tag
  class PageSetup < OOXMLDocumentObject
    # @return [Integer] id of paper size
    attr_reader :paper_size
    # @return [Symbol] orientation of page
    attr_reader :orientation

    # Parse PageSetup object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [PageSetup] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'paperSize'
          @paper_size = value.value.to_i
        when 'orientation'
          @orientation = value_to_symbol(value)
        end
      end
      self
    end
  end
end
