module OoxmlParser
  # Class for parsing `w:rStyle` object
  class RunStyle < OOXMLDocumentObject
    # @return [Integer] value of size
    attr_accessor :value

    # Parse RunStyle
    # @param [Nokogiri::XML:Node] node with RunStyle
    # @return [RunStyle] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'val'
          @value = value.value.to_s
        end
      end
      self
    end

    # @return [DocumentStyle] style which was referenced in RunStyle
    def referenced
      root_object.document_style_by_id(@value)
    end
  end
end
