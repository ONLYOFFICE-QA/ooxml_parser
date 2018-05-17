require_relative 'color/docx_pattern_fill'
module OoxmlParser
  # Color inside DOCX
  class DocxColor < OOXMLDocumentObject
    attr_accessor :type, :value, :stretching_type, :alpha

    # Parse DocxColor object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [DocxColor] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'blipFill'
          @type = :picture
          @value = DocxBlip.new(parent: self).parse(node_child)
          node_child.xpath('*').each do |fill_type_node_child|
            case fill_type_node_child.name
            when 'tile'
              @stretching_type = :tile
            when 'stretch'
              @stretching_type = :stretch
            when 'blip'
              fill_type_node_child.xpath('alphaModFix').each { |alpha_node| @alpha = alpha_node.attribute('amt').value.to_i / 1_000.0 }
            end
          end
        when 'solidFill'
          @type = :solid
          @value = Color.new(parent: self).parse_color_model(node_child)
        when 'gradFill'
          @type = :gradient
          @value = GradientColor.new(parent: self).parse(node_child)
        when 'pattFill'
          @type = :pattern
          @value = DocxPatternFill.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
