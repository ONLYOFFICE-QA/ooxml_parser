# frozen_string_literal: true

module OoxmlParser
  # Class for `pane` data
  class Pane < OOXMLDocumentObject
    attr_accessor :state, :top_left_cell, :x_split, :y_split

    # Parse Pane object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Pane] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'state'
          @state = value.value.to_sym
        when 'topLeftCell'
          @top_left_cell = Coordinates.parse_coordinates_from_string(value.value)
        when 'xSplit'
          @x_split = value.value
        when 'ySplit'
          @y_split = value.value
        end
      end
      self
    end
  end
end
