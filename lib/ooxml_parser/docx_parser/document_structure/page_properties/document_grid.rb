# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `docGrid` tags
  class DocumentGrid < OOXMLDocumentObject
    attr_accessor :type, :line_pitch, :char_space

    # Parse DocumentGrid
    # @param [Nokogiri::XML:Element] node with DocumentGrid
    # @return [DocumentGrid] value of DocumentGrid
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'charSpace'
          @char_space = value.value
        when 'linePitch'
          @line_pitch = value.value.to_i
        when 'type'
          @type = value.value
        end
      end
      self
    end
  end
end
