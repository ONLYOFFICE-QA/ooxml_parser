# DOCX Color Scheme
module OoxmlParser
  class DocxColorScheme < OOXMLDocumentObject
    attr_accessor :color, :type

    def initialize
      @color = Color.new
      @type = :unknown
    end

    def self.parse(color_scheme_node)
      color_scheme = DocxColorScheme.new
      color_scheme_node.xpath('*').each do |color_scheme_node_child|
        case color_scheme_node_child.name
        when 'solidFill'
          color_scheme.type = :solid
          color_scheme.color = Color.parse_color_model(color_scheme_node_child)
        when 'gradFill'
          color_scheme.type = :gradient
          color_scheme.color = GradientColor.parse(color_scheme_node_child)
        when 'noFill'
          color_scheme.color = :none
          color_scheme.type = :none
        when 'srgbClr'
          color_scheme.color = Color.from_int16(color_scheme_node_child.attribute('val').value)
        when 'schemeClr'
          color_scheme.color = Color.parse_scheme_color(color_scheme_node_child)
        end
      end
      color_scheme
    end

    def to_s
      "Color: #{@color}, type: #{type}"
    end
  end
end
