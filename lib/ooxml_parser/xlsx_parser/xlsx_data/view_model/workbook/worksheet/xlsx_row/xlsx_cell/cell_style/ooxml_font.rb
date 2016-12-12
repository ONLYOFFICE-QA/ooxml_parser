module OoxmlParser
  # Font Data
  class OOXMLFont < OOXMLDocumentObject
    attr_accessor :name, :size, :font_style, :color

    def initialize(parent: nil)
      @name = 'Calibri'
      @size = 11
      @parent = parent
    end

    def parse(style_number)
      @font_style = FontStyle.new
      font_style_node = XLSXWorkbook.styles_node.xpath('//xmlns:font')[style_number.to_i]
      font_style_node.xpath('*').each do |node_child|
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
          @color = Color.parse_color_tag(node_child)
        end
      end
      self
    end
  end
end
