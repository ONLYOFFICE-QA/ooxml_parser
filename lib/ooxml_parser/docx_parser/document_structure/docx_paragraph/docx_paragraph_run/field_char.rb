# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `w:fldChar` object
  class FieldChar < OOXMLDocumentObject
    # @return [Symbol] type of field char
    attr_accessor :type

    # Parse FieldChar
    # @param [Nokogiri::XML:Node] node with FieldChar
    # @return [FieldChar] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'fldCharType'
          @type = value.value.to_sym
        end
      end
      self
    end
  end
end
