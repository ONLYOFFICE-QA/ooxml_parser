# frozen_string_literal: true

module OoxmlParser
  # Class for parsing <item> tag
  class Item < OOXMLDocumentObject
    # @return [Integer] index of item
    attr_reader :index
    # @return [Symbol] type of item
    attr_reader :type

    # Parse `<item>` tag
    # # @param [Nokogiri::XML:Element] node with Item data
    # @return [item]
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'x'
          @index = value.value.to_i
        when 't'
          @type = value.value.to_sym
        end
      end
      self
    end
  end
end
