# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `w:pgBorders` PageBorders object
  class PageBorders < OOXMLDocumentObject
    # @return [Symbol] display value
    attr_reader :display
    # @return [Symbol] offset from value
    attr_reader :offset_from

    # Parse PageBorders
    # @param [Nokogiri::XML:Node] node with PageBorders
    # @return [PageBorders] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'display'
          @display = value.value.to_sym
        when 'offsetFrom'
          @offset_from = value.value.to_sym
        end
      end
      self
    end
  end
end
