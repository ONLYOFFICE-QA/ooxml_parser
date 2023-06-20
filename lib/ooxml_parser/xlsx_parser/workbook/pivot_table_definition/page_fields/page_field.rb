# frozen_string_literal: true

module OoxmlParser
  # Class for parsing <pageField> tag
  class PageField < OOXMLDocumentObject
    # @return [Integer] index of the field
    attr_accessor :field
    # @return [Integer]  index of the OLAP hierarchy
    attr_reader :hierarchy

    # Parse `<pageField>` tag
    # # @param [Nokogiri::XML:Element] node with PageField data
    # @return [PageField]
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'fld'
          @field = value.value.to_i
        when 'hier'
          @hierarchy = value.value.to_i
        end
      end
      self
    end
  end
end
