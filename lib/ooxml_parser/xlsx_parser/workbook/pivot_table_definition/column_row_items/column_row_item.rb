# frozen_string_literal: true

module OoxmlParser
  # Class for parsing <i> tag
  class ColumnRowItem < OOXMLDocumentObject
    # @return [Symbol] type of item
    attr_reader :type

    # Parse `<location>` tag
    # @param [Nokogiri::XML:Element] node with location data
    # @return [Location]
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 't'
          @type = value.value.to_sym
        end
      end
      self
    end
  end
end
