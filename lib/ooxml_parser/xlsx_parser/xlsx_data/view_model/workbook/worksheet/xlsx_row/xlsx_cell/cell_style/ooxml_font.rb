# Font Data
module OoxmlParser
  class OOXMLFont < OOXMLDocumentObject
    attr_accessor :name, :size, :font_style, :color

    def initialize(name = 'Calibri', size = '11', font_style = nil, color = nil)
      @name = name
      @size = size
      @font_style = font_style
      @color = color
    end

    def self.parse(style_number)
      font = OOXMLFont.new
      font_style_node = XLSXWorkbook.styles_node.xpath('//xmlns:font')[style_number.to_i]
      font.name = font_style_node.xpath('xmlns:name').first.attribute('val').value
      font.size = font_style_node.xpath('xmlns:sz').first.attribute('val').value.to_i if font_style_node.xpath('xmlns:sz').first
      font.font_style = FontStyle.new
      font_style_node.xpath('*').each do |font_style_node_child|
        case font_style_node_child.name
        when 'b'
          font.font_style.bold = true
        when 'i'
          font.font_style.italic = true
        when 'strike'
          font.font_style.strike = :single
        when 'u'
          font.font_style.underlined = Underline.new(:single)
        when 'color'
          font.color = Color.parse_color_tag(font_style_node_child)
        end
      end
      font
    end
  end
end
