# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `m:grow` object
  class NaryGrow < OOXMLDocumentObject
    # @return [String] value of grow
    attr_accessor :value

    # Parse NaryGrow
    # @param [Nokogiri::XML:Node] node with NaryGrow
    # @return [NaryGrow] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'val'
          @value = value.value
        end
      end
      self
    end
  end
end
