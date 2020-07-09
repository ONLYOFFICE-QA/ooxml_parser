# frozen_string_literal: true

module OoxmlParser
  # Parsing `fonts` tag
  class Font < OOXMLDocumentObject
    attr_accessor :name, :size, :font_style, :color
    # @return [ValuedChild] vertical alignment of font
    attr_reader :vertical_alignment

    def initialize(parent: nil)
      @name = 'Calibri'
      @size = 11
      @parent = parent
    end

    # Parse Font object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Font] result of parsing
    def parse(node)
      @font_style = FontStyle.new
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'name'
          @name = node_child.attribute('val').value
        when 'sz'
          @size = node_child.attribute('val').value.to_f
        when 'b'
          @font_style.bold = true
        when 'i'
          @font_style.italic = true
        when 'strike'
          @font_style.strike = :single
        when 'u'
          @font_style.underlined = Underline.new(:single)
        when 'color'
          @color = OoxmlColor.new(parent: self).parse(node_child)
        when 'vertAlign'
          @vertical_alignment = ValuedChild.new(:symbol, parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
