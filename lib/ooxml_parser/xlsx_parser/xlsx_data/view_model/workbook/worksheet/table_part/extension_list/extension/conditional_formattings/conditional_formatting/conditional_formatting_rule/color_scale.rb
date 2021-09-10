# frozen_string_literal: true

module OoxmlParser
  # Class for `colorScale` data
  class ColorScale < OOXMLDocumentObject
    # @return [Array, ConditionalFormatValueObject] list of values
    attr_reader :values
    # @return [Array, Color] list of colors
    attr_reader :colors

    def initialize(parent: nil)
      @values = []
      @colors = []
      super
    end

    # Parse ColorScale data
    # @param [Nokogiri::XML:Element] node with ColorScale data
    # @return [ColorScale] value of ColorScale data
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'cfvo'
          @values << ConditionalFormatValueObject.new(parent: self).parse(node_child)
        when 'color'
          @colors << OoxmlColor.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
