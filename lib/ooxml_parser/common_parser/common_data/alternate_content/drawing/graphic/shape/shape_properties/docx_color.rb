require_relative 'color/gradient_color/docx_gradient_color'
require_relative 'color/docx_pattern_fill'
# Color inside DOCX
module OoxmlParser
  class DocxColor < OOXMLDocumentObject
    attr_accessor :type, :value, :stretching_type, :alpha

    def initialize(value = nil)
      @value = value
    end

    def self.parse(shape_properties_node)
      fill_color = DocxColor.new
      shape_properties_node.xpath('*').each do |fill_node|
        case fill_node.name
        when 'blipFill'
          fill_color.type = :picture
          fill_color.value = DocxBlip.parse(fill_node)
          fill_node.xpath('*').each do |fill_type_node_child|
            case fill_type_node_child.name
            when 'tile'
              fill_color.stretching_type = :tile
            when 'stretch'
              fill_color.stretching_type = :stretch
            when 'blip'
              fill_type_node_child.xpath('alphaModFix').each { |alpha_node| fill_color.alpha = alpha_node.attribute('amt').value.to_i / 1_000.0 }
            end
          end
        when 'solidFill'
          fill_color.type = :solid
          fill_color.value = Color.parse_color_model(fill_node)
        when 'gradFill'
          fill_color.type = :gradient
          fill_color.value = GradientColor.parse(fill_node)
        when 'pattFill'
          fill_color.type = :pattern
          fill_color.value = DocxPatternFill.parse(fill_node)
        end
      end
      fill_color
    end
  end
end
