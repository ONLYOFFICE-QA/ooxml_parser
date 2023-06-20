# frozen_string_literal: true

module OoxmlParser
  # Class for parsing <field> tag
  class Field < OOXMLDocumentObject
    # @return [Integer] the index to a pivotField item value
    attr_accessor :field_index

    # Parse `<field>` tag
    # # @param [Nokogiri::XML:Element] node with Field data
    # @return [Field]
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'x'
          @field_index = value.value.to_i
        end
      end
      self
    end
  end
end
