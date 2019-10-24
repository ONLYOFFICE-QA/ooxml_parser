# frozen_string_literal: true

module OoxmlParser
  # Fallback DOCX shape fill properties
  class OldDocxShapeFill < OOXMLDocumentObject
    attr_accessor :stretching_type, :opacity, :title
    # @return [FileReference] image structure
    attr_accessor :file_reference

    # Parse OldDocxShapeFill object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [OldDocxShapeFill] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'id'
          @file_reference = FileReference.new(parent: self).parse(node)
        when 'type'
          @stretching_type = case value.value
                             when 'frame'
                               :stretch
                             else
                               value.value.to_sym
                             end
        when 'title'
          @title = value.value.to_s
        end
      end
      self
    end
  end
end
