# frozen_string_literal: true

module OoxmlParser
  # Chart Style
  class ChartStyle < OOXMLDocumentObject
    attr_accessor :style_number
    # @return [OleObject] ole object
    attr_accessor :ole_object

    # Parse ChartStyle
    # @param [Nokogiri::XML:Node] node with Relationships
    # @return [ChartStyle] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'style'
          @style_number = node_child.attribute('val').value.to_i
        when 'oleObject'
          @ole_object = OleObject.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
