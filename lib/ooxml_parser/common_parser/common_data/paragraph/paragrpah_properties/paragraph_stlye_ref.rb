# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `pStyle` tags
  class ParagraphStyleRef < OOXMLDocumentObject
    # @return [Integer] value of ParagraphStyleRef
    attr_reader :value

    # Parse ParagraphStyleRef object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [ParagraphStyleRef] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'val'
          @value = value.value.to_s
        end
      end
      self
    end

    # @return [ParagraphStyle] which was referenced
    def referenced_style
      root_object.document_style_by_id(value)
    end
  end
end
