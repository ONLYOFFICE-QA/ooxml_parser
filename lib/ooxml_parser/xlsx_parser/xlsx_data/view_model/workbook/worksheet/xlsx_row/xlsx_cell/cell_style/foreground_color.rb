# Foreground Color Data
module OoxmlParser
  class ForegroundColor < OOXMLDocumentObject
    attr_accessor :type, :color

    def initialize(type = nil, color = Color.new)
      @type = type
      @color = color
    end

    def nil?
      @type.nil? && @color.nil?
    end

    def self.parse(style_number)
      fill_color = ForegroundColor.new
      fill_style_node = XLSXWorkbook.styles_node.xpath('//xmlns:fill')[style_number.to_i].xpath('xmlns:patternFill')
      if fill_style_node && !fill_style_node.empty?
        fill_color.type = fill_style_node.attribute('patternType').value.to_sym
        fill_color.color = Color.parse_color_tag(fill_style_node.xpath('xmlns:fgColor').first)
      end
      fill_color
    end
  end
end
