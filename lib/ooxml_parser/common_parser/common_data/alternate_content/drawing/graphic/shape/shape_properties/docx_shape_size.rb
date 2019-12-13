# frozen_string_literal: true

module OoxmlParser
  # Docx Shape Size
  class DocxShapeSize < OOXMLDocumentObject
    attr_accessor :rotation, :offset, :extent
    attr_accessor :child_offset
    attr_accessor :child_extent
    # @return [True, False] is image flipped horizontally
    attr_reader :flip_horizontal
    # @return [True, False] is image flipped vertically
    attr_reader :flip_vertical

    alias extents extent

    # Parse DocxShapeSize object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [DocxShapeSize] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'rot'
          @rotation = OoxmlSize.new(value.value.to_f, :one_60000th_degree)
        when 'flipH'
          @flip_horizontal = attribute_enabled?(value)
        when 'flipV'
          @flip_vertical = attribute_enabled?(value)
        end
      end
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'off'
          @offset = OOXMLCoordinates.parse(node_child, unit: :emu)
        when 'ext'
          @extent = OOXMLCoordinates.parse(node_child, x_attr: 'cx', y_attr: 'cy', unit: :emu)
        when 'chOff'
          @child_offset = OOXMLCoordinates.parse(node_child, unit: :emu)
        when 'chExt'
          @child_extent = OOXMLCoordinates.parse(node_child, x_attr: 'cx', y_attr: 'cy', unit: :emu)
        end
      end
      self
    end
  end
end
