# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `m:argSz` object
  class ArgumentSize < OOXMLDocumentObject
    # @return [String] value of size
    attr_accessor :value

    # Parse ArgumentSize
    # @param [Nokogiri::XML:Node] node with ArgumentSize
    # @return [ArgumentSize] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'val'
          @value = value.value.to_f
        end
      end
      self
    end
  end
end
