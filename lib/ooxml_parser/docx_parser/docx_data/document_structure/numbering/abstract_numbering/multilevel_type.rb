# frozen_string_literal: true

module OoxmlParser
  # This element specifies the type of numbering defined by a given abstract numbering type
  class MultilevelType < OOXMLDocumentObject
    # @return [String] value of Multilevel Type
    attr_accessor :value

    # Parse MultilevelType
    # @param [Nokogiri::XML:Node] node with MultilevelType
    # @return [MultilevelType] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'val'
          @value = value_to_symbol(value)
        end
      end
      self
    end
  end
end
