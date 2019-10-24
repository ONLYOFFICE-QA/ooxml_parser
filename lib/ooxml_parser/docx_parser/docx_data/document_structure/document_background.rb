# frozen_string_literal: true

module OoxmlParser
  # Class for describing Document Background `w:background`
  class DocumentBackground < OOXMLDocumentObject
    attr_reader :color1, :size, :color2, :type
    # @return [FileReference] image structure
    attr_reader :file_reference
    # @return [Fill] fill data
    attr_reader :fill

    def initialize(parent: nil)
      @color1 = nil
      @type = 'simple'
      @parent = parent
    end

    # Parse DocumentBackground object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [DocumentBackground] result of parsing
    def parse(node)
      @color1 = Color.new(parent: self).parse_hex_string(node.attribute('color').value)
      node.xpath('v:background').each do |second_color|
        @size = second_color.attribute('targetscreensize').value.sub(',', 'x') unless second_color.attribute('targetscreensize').nil?
        second_color.xpath('*').each do |node_child|
          case node_child.name
          when 'fill'
            @fill = Fill.new(parent: self).parse(node_child)
          end
        end
      end
      self
    end
  end
end
