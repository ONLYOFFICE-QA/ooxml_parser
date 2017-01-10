module OoxmlParser
  # Class for parsing `sldSz` tag
  class SlideSize < OOXMLDocumentObject
    attr_accessor :width, :height, :type

    # Parse SlideSize object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Bookmark] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'cx'
          @width = OoxmlSize.new(value.value.to_f, :emu)
        when 'cy'
          @height = OoxmlSize.new(value.value.to_f, :emu)
        when 'type'
          @type = value.value.to_sym
        end
      end
      self
    end
  end
end
