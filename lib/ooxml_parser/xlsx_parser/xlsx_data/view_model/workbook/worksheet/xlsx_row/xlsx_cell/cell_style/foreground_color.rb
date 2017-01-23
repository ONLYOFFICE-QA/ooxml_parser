module OoxmlParser
  # Foreground Color Data
  class ForegroundColor < OOXMLDocumentObject
    attr_accessor :type, :color

    def initialize(type = nil, color = Color.new, parent: nil)
      @type = type
      @color = color
      @parent = parent
    end

    def nil?
      @type.nil? && @color.nil?
    end

    # Parse ForegroundColor object
    # @param style_number [Integer) number of style
    # @return [ForegroundColor] result of parsing
    def parse(style_number)
      fill_style_node = XLSXWorkbook.styles_node.xpath('//xmlns:fill')[style_number.to_i].xpath('xmlns:patternFill')
      if fill_style_node && !fill_style_node.empty?
        @type = fill_style_node.attribute('patternType').value.to_sym if fill_style_node.attribute('patternType')
        @color = Color.parse_color_tag(fill_style_node.xpath('xmlns:fgColor').first) if fill_style_node.xpath('xmlns:fgColor')
      end
      self
    end
  end
end
