# frozen_string_literal: true

module OoxmlParser
  # DOCX Color Scheme
  class DocxColorScheme < OOXMLDocumentObject
    attr_accessor :color, :type

    def initialize(parent: nil)
      @color = Color.new
      @type = :unknown
      super
    end

    # Parse DocxColorScheme object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [DocxColorScheme] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'solidFill'
          @type = :solid
          @color = Color.new(parent: self).parse_color_model(node_child)
        when 'gradFill'
          @type = :gradient
          @color = GradientColor.new(parent: self).parse(node_child)
        when 'noFill'
          @color = :none
          @type = :none
        when 'srgbClr'
          @color = Color.new(parent: self).parse_hex_string(node_child.attribute('val').value)
        when 'schemeClr'
          @color = Color.new(parent: self).parse_scheme_color(node_child)
        end
      end
      self
    end

    # @return [String] result of convert of object to string
    def to_s
      "Color: #{@color}, type: #{type}"
    end
  end
end
